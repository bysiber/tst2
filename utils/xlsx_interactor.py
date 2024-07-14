from collections import defaultdict
import copy
import os
import pandas as pd
from dto.configurations import AccountInformation
from _utils.logger import logger


class ExcelHelper:

    def read_xlsx(self, file_location: str, file_name: str, sheet_name: str):
        return pd.read_excel(os.path.join(file_location, file_name), sheet_name)

    def extract_information_from_file_to_class(self, file: pd.DataFrame):
        rows = []
        for _, row in file.iterrows():
            transaction = AccountInformation(row["Company Number"], row["Card Account Number"], row["CH Full Name"],
                                    row["Card Embossed Line 1"], row["Card Embossed Line 2"], row["CH Email"],
                                    row["Grp Full Name"], row["Card Accounting Code"])
            rows.append(transaction)
        
        return rows
    
    def complete_8_digits_card_numbers(self, file_from: pd.DataFrame, file_to: pd.DataFrame):
        account_info_to_number = defaultdict(list)
        card_numbers_not_found = []
        
        file_from['Card Account Number'] = file_from['Card Account Number'].str.replace('x', 'X')
        
        for _, row_from in file_from.iterrows():
            key = (row_from["Card Account Number"][-4:]), row_from["CH Full Name"].upper(), row_from["Card Embossed Line 2"].upper()
            account_info_to_number[key].append(row_from["Card Account Number"])
        
        for i, row_to in file_to.iterrows():
            key = (row_to["Card Account Number"][-4:]), row_to["CH Full Name"].upper(), row_to["Card Embossed Line 2"].upper()
            if key in account_info_to_number:
                file_to.at[i, "Card Account Number"] = account_info_to_number[key][0]
            else:
                card_numbers_not_found.append(file_to.at[i, "Card Account Number"])
                
        if len(card_numbers_not_found) != 0:
            logger.warning(f"Update last 8 of CC as the following cards were not found: {card_numbers_not_found}")
            logger.warning(f"This is a total of {len(card_numbers_not_found)} unmatched rows from MasterFile (Cards).")

            
    def create_file_filters_and_tabs(self, file_from: pd.DataFrame, file_to: pd.DataFrame, tabs_dict: dict):
        files_to_create = file_from['File'].unique()
        for file in files_to_create:
            tabs_to_create = file_from.loc[file_from['File'] == file, 'Tab']

            for tab in tabs_to_create:

                if file == "Suspended Hotels Report":
                    suspended_rows = file_to[file_to['Grp Full Name'].str.contains('suspend', case=False)]
                    if not suspended_rows.empty and len(suspended_rows) > 0:
                        tabs_dict[(file, tab)] = suspended_rows
                    continue

                is_complement_tab = file_from.loc[file_from['Tab'] == tab, 'Complement'].iloc[0]
        
                if is_complement_tab == 'Yes':
                    self.create_file_filters_and_tabs_based_on_difference((file, "Full EOM Report FS&SS"), 
                        (file, "Select Service"), tabs_dict)
                else:
                    filters = file_from.loc[file_from['Tab'] == tab, 'Filters'].iloc[0]
                    
                    if filters is None or pd.isna(filters):
                        tabs_dict[(file, tab)] = file_to
                        continue
                    
                    filters_names = self.remove_extra_spaces_from_list(filters.split(';'))
                    filter_iterator = 0
                    
                    for filter in filters_names:
                        if filter in file_from.columns:
                            filter_values = file_from.loc[file_from['Tab'] == tab, filter].unique()
                            
                            if len(filter_values) > 0:
                                filter_values = str(filter_values[0])
                            else:
                                logger.error(f"There are no values configured to filter in the column: '{filter}' for the tab: '{tab}'")
                                raise ValueError(f"There are no values configured to filter in the column: '{filter}' for the tab: '{tab}'")
                                
                            filter_values_names = self.remove_extra_spaces_from_list(filter_values.split(';'))
                            exclude_filter = self.get_value_at_position(file_from.loc[(file_from['Tab'] == tab), 'Exclude'].iloc[0], filter_iterator)
                            
                            if len(filter_values_names) > 0:                            
                                final_tab_information = file_to[file_to[filter].isin(filter_values_names)]
                                
                                if exclude_filter == 'Yes':
                                    if len(filter_values) > 0:
                                        items_to_remove = file_to[file_to[filter].isin(filter_values_names)]

                                        final_tab_information = file_to[~tabs_dict[(file, tab)].isin(items_to_remove)].drop_duplicates().dropna(how='all')
                                    else:
                                        logger.error(f"There are no values (TO EXCLUDE) configured to filter in the column: '{filter}' for the tab: '{tab}'")
                                        raise ValueError(f"There are no values (TO EXCLUDE) configured to filter in the column: '{filter}' for the tab: '{tab}'")
                                
                                previous_value = tabs_dict.get((file, tab))
                                if previous_value is not None:
                                    merged_df = pd.concat([previous_value, final_tab_information])
                                    tabs_dict[(file, tab)] = merged_df
                                else:
                                    tabs_dict[(file, tab)] = final_tab_information
                            else:
                                logger.error(f"There are no columns configured to filter in the column: '{filter}' for the tab: '{tab}'")
                                raise ValueError(f"There are no columns configured to filter in the column: '{filter}' for the tab: '{tab}'")
                        else:
                            logger.error(f"There is no column that details the filters for this tab: '{tab}'")
                            raise ValueError(f"There is no column that details the filters for this tab: '{tab}'")
                        
                        filter_iterator += 1
                    
    def create_file_filters_and_tabs_based_on_difference(self, tab_name1: (str, str), tab_name2: (str, str), tabs_dict: dict):
        full_report_tab = tabs_dict[tab_name1]
        select_service_tab = tabs_dict[tab_name2]

        if full_report_tab is not None and select_service_tab is not None:
            
            select_service_tab = select_service_tab.drop_duplicates()
            full_report_tab = full_report_tab.drop_duplicates()
            difference = full_report_tab[~full_report_tab.isin(select_service_tab)].dropna(how='all')
                        
            new_file_name = tab_name1[0]
            new_tab_name = "Full Service"
            
            filtered_difference = difference[~difference['Grp Full Name'].str.contains('suspend', case=False)]

            tabs_dict[(new_file_name, new_tab_name)] = filtered_difference
        else:
            logger.error(f"'{full_report_tab}' tab or '{select_service_tab}' tab is not existing in the file: '{new_file_name}'")
            raise ValueError(f"'{full_report_tab}' tab or '{select_service_tab}' tab is not existing in the file: '{new_file_name}'")
                
    def create_remaining_tabs_fs(self, file_name: str, tab_name: str, tabs_dict: dict):
        full_service_tab = None
        for (file, tab), data_frame in tabs_dict.items():
            if tab_name in tab:
                full_service_tab = data_frame
                
        if full_service_tab is not None:
            unique_values = full_service_tab['Card Embossed Line 2'].unique()

            for value in unique_values:
                filtered_dfs = full_service_tab.loc[full_service_tab['Card Embossed Line 2'] == value]
                tabs_dict[(file_name, self.camelize(value))] = filtered_dfs
        else:
            logger.error(f"The tab configured for: '{tab_name}' is not existing in the file: '{file}'")
            raise ValueError(f"The tab configured for: '{tab_name}' is not existing in the file: '{file}'")
        
    def create_remaining_tabs_ss(self, select_service_tab: pd.DataFrame, tabs_dict_to: dict):
        if select_service_tab is not None:
            unique_values = select_service_tab['Card Embossed Line 2'].unique()
            new_file_name = "Select Service EOM Report"

            for value in unique_values:
                filtered_dfs = select_service_tab.loc[select_service_tab['Card Embossed Line 2'] == value]
                tabs_dict_to[(new_file_name, self.camelize(value))] = filtered_dfs
        else:
            logger.error(f"The tab configured for: '{select_service_tab}' is not existing in the file: 'Select Service EOM Report'")
            raise ValueError(f"The tab configured for: '{select_service_tab}' is not existing in the file: 'Select Service EOM Report'")

    def create_payment_file(self, tab_name: str, payments_tab: pd.DataFrame, tabs_dict_to: dict):
        if not tabs_dict_to:
            new_file_name = "FS Hotel"
            tabs_dict_to[(new_file_name, tab_name)] = payments_tab
        else:
            logger.error(f"{new_file_name} is not empty. '{tab_name}' tab was not created.")
            raise ValueError(f"{new_file_name} is not empty. '{tab_name}' tab was not created.")

    def complete_specific_files(self, dict_from: dict, tabs_dict_fs: dict, tabs_dict_sh: dict):
        fs_was_created = False
        ss_was_created = False
        sh_was_created = False
        dict_from_copy = copy.deepcopy(dict_from)
        
        for (file, tab), data_frame in dict_from.items():
            if "Full Service" in tab:
                file_name = "Full Service EOM Report"
                tabs_dict_fs[(file_name, "ALL FS")] = data_frame
                self.create_remaining_tabs_fs(file_name, "ALL FS", tabs_dict_fs)
                fs_was_created = True
                
            elif "Select Service EOM Report" in file and "Colony" in tab:
                self.create_remaining_tabs_ss(data_frame, dict_from_copy)
                
                tab_to_move = "Corepoint"
                corepoint_data = dict_from_copy[(file, tab_to_move)]
                key_to_remove = (file, tab_to_move)
                del dict_from_copy[key_to_remove]
                dict_from_copy[(file, tab_to_move)] = corepoint_data
                    
                ss_was_created = True

            elif "Suspended Hotels Report" in file and "Suspended Hotels" in tab:
                tabs_dict_sh[(file, tab)] = data_frame
                sh_was_created = True
                
        if sh_was_created:
            logger.warning(f"Suspended Hotels detected! {file} file and {tab} tab were created.")
        else:
            logger.warning(f"No suspended hotels were detected.")

        if not fs_was_created or not ss_was_created:
            logger.error(f"Full Service or Select Service reports were not created. Please validate the information from tab: '{tab}' in '{file}'.")
            raise ValueError(f"Full Service or Select Service reports were not created. Please validate the information from tab: '{tab}' in '{file}'.")

    def complete_totals_from_files(self, tabs_list: list[dict]):
        for tab_dict in tabs_list:
            for (file, tab), data_frame in tab_dict.items():
                if data_frame is not None and not data_frame.empty:
                    data_frame_copy = data_frame.copy()
                    total_amount_for_tab = data_frame_copy['Item Total'].sum()

                    if 'Item Total' not in data_frame_copy.columns:
                        logger.warning(
                            f"'Item Total' column not found in the data frame for tab '{tab}' in '{file}'. Skipping.")
                        continue
                    if (file == "EOM Pcard Report - FS & SS") and (tab == "Full Service"):
                        num_cols = data_frame_copy.shape[1]
                        col_index_to_set = data_frame_copy.columns.get_loc("Item Total")
                        new_row_data = [None] * num_cols
                        new_row_data[col_index_to_set] = total_amount_for_tab
                        new_row = pd.DataFrame([new_row_data], columns=data_frame_copy.columns)
                        data_frame_copy = pd.concat([data_frame_copy, new_row], ignore_index=True)
                        logger.info(f"Last row Full Service: {data_frame_copy.tail(1)}")
                    else:
                        data_frame_copy.loc[len(data_frame_copy) + 1, 'Item Total'] = total_amount_for_tab
                    tab_dict[(file, tab)] = data_frame_copy
                else:
                    logger.info(
                        f"Not able to generate the sum of 'Item Total' for the following tab: '{tab}' in '{file}'.")

    def complete_subtotal_if_match(self, total_amount: int, payment_tab: dict, hotel_tab: dict, tab_name: str):
        if total_amount != None and total_amount != 0:
            hotel_names = payment_tab['Full Service'].tolist()

            payment_tab['Full Service'] = payment_tab['Full Service'].str.title()
            hotel_names = [name.title() for name in hotel_names]
            tab_name = tab_name.title()

            if tab_name in hotel_names:
                mask = payment_tab['Full Service'] == tab_name
                payment_tab.loc[mask, 'B/U'] = total_amount
                return True
            else:
                return False
        else:
            logger.error(f"The tab configured for: '{tab_name}' has total a total amount = 0. Process will continue executing the next tab.")

    def create_total_sum_for_hotels(self, payment_tab: dict):

        row_to_write = len(payment_tab.index) + 1
        payment_tab.loc[row_to_write] = None
        row_to_write += 1
        
        payment_tab.loc[row_to_write, 'ACCT#'] = 'TOTAL AMOUNT:'
        payment_tab.loc[row_to_write, 'B/U'] = payment_tab['B/U'].sum(skipna=True)

    def get_data_frame_from_filter(self, file: pd.DataFrame, column: str, filters: list, is_in: str):
        if is_in == 'N':
            return file[file[column].isin(filters)]
        elif is_in == 'Y':
            return file[~file[column].isin(filters)]
        else:
            return None
    
    def remove_extra_spaces_from_list(self, texts: list):
        for i in range(len(texts)):
            texts[i] = texts[i].strip()
        return texts
    
    def remove_extra_spaces(self, text: str):
        return text.strip()
    
    def camelize(self, text: str):
        words = text.split()
        return ' '.join(word.capitalize() for word in words)

    def get_value_at_position(self, string: str, position: int):
        values = [value.strip() for value in string.split(',')]

        if 0 <= position < len(values):
            return values[position]
        else:
            return None