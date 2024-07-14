class AccountInformation:
    
    def __init__(self, company_number: int, card_account_number: int, ch_full_name: str, card_embossed_line_1: str, 
                 card_embossed_line_2: str, ch_email: str, grp_full_name: str, card_accounting_code: str):
        
        self.company_number = company_number
        self.card_account_number = card_account_number
        self.ch_full_name = ch_full_name
        self.card_embossed_line_1 = card_embossed_line_1
        self.card_embossed_line_2 = card_embossed_line_2
        self.ch_email = ch_email
        self.grp_full_name = grp_full_name
        self.card_accounting_code = card_accounting_code
        
        
class ResumedAccountInformation:
    
    def __init__(self, card_account_number: int, ch_full_name: str, card_embossed_line_2: str):
        
        self.card_account_number = card_account_number
        self.ch_full_name = ch_full_name
        self.card_embossed_line_2 = card_embossed_line_2
        
    
    def matches(self, resumed_account_information_to_compare):
        if (str(self.card_account_number)[-4:] == str(resumed_account_information_to_compare.card_account_number)[-4:]
            and self.ch_full_name.upper() == resumed_account_information_to_compare.ch_full_name.upper()
            and self.card_embossed_line_2.upper() == resumed_account_information_to_compare.card_embossed_line_2.upper()):
            return True
        else:
            return False