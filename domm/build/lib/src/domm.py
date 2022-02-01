"""
Email domain name matching for generator contact
"""
import pandas as pd
from os.path import exists
import sys

# Column titles for parsed data
NUM_SITE_WO = "Number_Sites_WO"
SITE_ID_WO = "Site_IDs_WO"
NUM_MAN_WO = "Number_Manifest_WO"
USER_EMAIL = "User_Emails"

# Raw data column titles
RAW_CON_EMAIL = "WITHOUT_CONTACT_EMAIL"
RAW_SM_EMAIL = "WITH_MANAGER_EMAILS"
RAW_NUM_MAN = "WITHOUT_NUM_MANIFESTS"
RAW_GEN_ID = "HANDLER_ID"

# common domain strings
COM_DOM = [".com", ".org", ".io", ".edu", ".mil", ".gov", ".net", ".us", ".biz"]
PUB_DOM = ["HOTMAIL.COM", "GMAIL.COM", "YAHOO.COM", "AOL.COM", "ATT.NET",
           "VERIZON.NET", "OUTLOOK.COM", "OUTLOOKGROUP.COM", "YAHOO.CO"]


def read_raw_data(args):
    """ read site info from excel/csv file"""
    print("opening", args.path)
    if args.path[-3:] == "csv":
        raw_data = pd.read_csv(args.path, keep_default_na=False, na_values=[''])
    elif args.path[-4:] == "xlsx" or args.path[-5:-1] == "xlsm":
        if args.sheet:
            raw_data = pd.read_excel(args.path, sheet_name=args.sheet, keep_default_na=False)
        else:
            raw_data = pd.read_excel(args.path, keep_default_na=False)
    print("File read")
    # ToDo: add exception for non csv/xlsx files
    return raw_data


def group_sites(unique_emails, raw_data):
    """The meat and potatoes"""
    con_email_data = raw_data[RAW_CON_EMAIL]
    user_emails_data = raw_data[RAW_SM_EMAIL]
    gen_df = {NUM_SITE_WO: [0] * len(unique_emails),
              SITE_ID_WO: [""] * len(unique_emails),
              NUM_MAN_WO: [0] * len(unique_emails),
              USER_EMAIL: unique_emails,
              }
    gen_df = pd.DataFrame(gen_df)
    for n in range(0, len(unique_emails)):
        con = unique_emails[n]
        num_sites_wo = 0
        num_manifest_wo = 0
        site_list_wo = ""
        if n % 100 == 0:
            print("\r" + str(int((n / len(unique_emails)) * 100)) + "% ", end=" ")
        for i in range(0, len(raw_data[RAW_GEN_ID])):
            if con == raw_data[RAW_SM_EMAIL][i]:
                num_sites_wo += 1
                site_list_wo = site_list_wo + raw_data[RAW_GEN_ID][i] + "; "
                num_manifest_wo = num_manifest_wo + int(raw_data[RAW_NUM_MAN][i])
        gen_df.loc[n, NUM_SITE_WO] = num_sites_wo
        gen_df.loc[n, SITE_ID_WO] = site_list_wo
        gen_df.loc[n, NUM_MAN_WO] = num_manifest_wo
    return gen_df


def display_contacts(user_emails):
    for emails_string in user_emails:
        print(emails_string)


def get_unique_contacts(raw_data):
    print("reading user emails")
    unique_sm_emails = raw_data[RAW_SM_EMAIL].unique()
    return unique_sm_emails


def display_stats(data):
    print("under construction")


def clean_input(raw_args):
    if raw_args.output:
        if exists(raw_args.output):
            overwrite_file = input("File exists. Overwrite file? [y/n] ")
            print(overwrite_file)
            if overwrite_file != "y":
                print("Exiting")
                sys.exit(0)
    if raw_args.sheet:
        if raw_args.sheet.isnumeric(): 
            raw_args.sheet = int(raw_args.sheet)


# Run/Test
def run(args):
    clean_input(args)
    data = read_raw_data(args)
    unique_contacts = get_unique_contacts(data)
    if args.display == "contacts":
        display_contacts(unique_contacts)
    if args.output:
        gen_df = group_sites(unique_contacts, data)
        print("writing to output")
        gen_df.to_excel(args.output, header=True, index=False, verbose=False)
