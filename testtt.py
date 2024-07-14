from pathlib import Path
import os


class FileNames:
    rules_master_filename = Path('Masterfile-Rules.xlsx')

class Paths:
    resources_filepath = Path(os.getcwd(), 'boa_pcard_reporting/resources')
    rules_filepath = Path(resources_filepath, 'rules')
    downloads_filepath = Path(resources_filepath, 'downloads')
    processed_filepath = Path(resources_filepath, 'processed')
    
    sp_root = Path("02 - BoA PCard Reporting")
    sp_rules_master_filepath = Path(sp_root, "01 - Rules", FileNames.rules_master_filename)
    sp_processed_filepath = Path(sp_root, "02 - Reports")



#print all the paths
print(Paths.resources_filepath)
print(Paths.rules_filepath)
print(Paths.downloads_filepath)
print(Paths.processed_filepath)
print(Paths.sp_root)
print(Paths.sp_rules_master_filepath)
print(Paths.sp_processed_filepath)
"""
# check the folder if there is no folder called that path create it
for path in [Paths.resources_filepath, 
             Paths.rules_filepath,
             Paths.downloads_filepath,
             Paths.processed_filepath]:
    if not os.path.exists(path):
        os.makedirs(path)
        print(f'Directory {path} created locally')"""



#check "/Sources" (text) path recursivly and print all dirs and files names under it structly
"""def print_dir_contents(path):
    for root, dirs, files in os.walk(path):
        level = root.replace(path, '').count(os.sep)
        indent = ' ' * 4 * (level)
        print(f'{indent}{os.path.basename(root)}/')
        subindent = ' ' * 4 * (level + 1)
        for f in files:
            print(f'{subindent}{f}')

print_dir_contents("Sources")"""