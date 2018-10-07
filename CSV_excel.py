from pyexcel.cookbook import merge_all_to_a_book
# import pyexcel.ext.xlsx # no longer required if you use pyexcel >= 0.2.2
import glob
import argparse


import csv
import os


def preprocess(file):
    inputFileName = file
    outputFileName = os.path.splitext(inputFileName)[0] + "_modified.csv"

    with open(inputFileName, 'rb') as inFile, open(outputFileName, 'wb') as outfile:
        r = csv.reader(inFile)
        w = csv.writer(outfile)

        next(r, None)  # skip the first row from the reader, the old header
        # write new header
        w.writerow(["protocol","first_name","last_name","email_address","role","site_id","is_sso_user"])

        # copy the rest
        for row in r:
            w.writerow(row)

    return outputFileName


def convert_csv_to_xls(csv_file_folder, output_file):
    preprocess_file = preprocess(csv_file_folder)
    merge_all_to_a_book(glob.glob("%s" % preprocess_file), output_file)




if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='converting csv file to xls file')
    parser.add_argument("--csv_folder", "-cf", help="csv files folder path", required=True)
    parser.add_argument("--analysis_file", "-of", help="The final output analysis file",
                        default="output_file.xlsx")
    args = parser.parse_args()
    convert_csv_to_xls(args.csv_folder, args.analysis_file)
