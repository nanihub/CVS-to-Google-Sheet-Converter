#!/usr/bin/env python3

import argparse

class servicenowargs:
    def __init__(self):
        parser = argparse.ArgumentParser(description='Copy Servicenow incs to spreadsheet.')

        parser.add_argument('-N', '--newspreadsheet',
                            help="Create a New spread sheet, Provide email id as well along \
                                 with your new spreadsheet name",
                            type=str)

        parser.add_argument('-n', '--newsheet',
                            help="To write data to new sheet in existing spreadsheet every \
                                 time it runs,\
                                 Provide Spreadsheetid, Csv file and newsheet name",
                            type=str)
        parser.add_argument('-c', '--existingsheet',
                            help="To append data to existing sheet, Please provide existing sheetname, \
                                 spreadsheetid, and csv file name",
                            type=str)

        parser.add_argument('-r', '--replacesheet',
                            help="Over write the existingsheet, Please provide existing sheetname, \
                                 spreadsheetid, and csv file name",
                            type=str)

        parser.add_argument('-i', '--spreadsheetid',
                            help="Spreasheet id where you want to perform your operations",
                            type=str)
        parser.add_argument('-o', '--outputfile',
                            help="Csv file name which you want to export to sheets",
                           type=str)

        parser.add_argument('-e', '--email',
                            help="Provide email id along with spreadsheet name",
                            type=str)
        
        
        args = parser.parse_args()

        if args.newsheet:
            self.newsheet = args.newsheet
        else:
            self.newsheet = None

        if args.spreadsheetid:
            self.spreadsheetid = args.spreadsheetid
        else:
            self.spreadsheetid = None

        if args.outputfile:
            self.outputfile = args.outputfile
        else:
            self.outputfile = None

        if args.newspreadsheet:
            self.newspreadsheet = args.newspreadsheet
        else:
            self.newspreadsheet = None

        if args.email:
            self.email = args.email
        else:
            self.email = None

        if args.existingsheet:
            self.existingsheet = args.existingsheet
        else:
            self.existingsheet = None

        if args.replacesheet:
            self.replacesheet = args.replacesheet
        else:
            self.replacesheet = None
