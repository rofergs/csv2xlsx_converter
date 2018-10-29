# CSV2EXCEL Converter
# Description: A simple script which allows command line conversion of .csv files to the .xlsx format.
# Note: Must be in the same folder as the .csv, output .xlsx file to the same folder.
# Usage in CLI: >>> python csv2xlsx.py -i *CSV_FILE" -d *DELIMETER* -o *XLSX_FILE*

import argparse  # For CLI functionality
import xlsxwriter  # For writing to the .xlsx file format
import os  # For the os.path function allowing to check if a given file exists within the script's directory.
import datetime  # To allow for adding the date and time to the output file
import logging  # For logging functionality


# Converter class
class CSVConverter:

    def __init__(self, csvfile, csvdelim, xlfile):
        self.csvfile = csvfile  # The .csv file the user wants to convert
        self.csvdelim = csvdelim  # The delimiter used in the .csv file
        self.xlfile = xlfile  # The user input naming the .xlsx file created
        self.xlfilefull = self.xlfile
        self.mypath = os.path.dirname(__file__)  # Sets the current filepath
        self.currentdate = datetime.datetime.now()  # Sets the current date and time to be used in file output

    # Extensions checking function
    def extensioncheck(self):
        logging.info("Start of extensioncheck()")
        # If the .csv file does not end with the .csv extension
        if not self.csvfile.lower().endswith(".csv"):
            # Add the .csv extension
            self.csvfile += ".csv"
            # DEBUG
            logging.debug("Added .csv extension as it was not present in input")
        # If the .csv file does end with the .csv extension
        else:
            # DEBUG
            logging.debug("Did not add .csv extensions as it was present in input ")

        # If the .xlsx file ends with the .xlsx extension
        if self.xlfile.lower().endswith(".xlsx"):
            # Split filename from extension, select the first object in the tuple and add timestamp and extension
            self.xlfilefull = os.path.splitext(self.xlfile)[0] + self.currentdate.strftime(" %d-%m-%Y") + ".xlsx"
            # DEBUG
            logging.debug("Added timestamp before .xlsx extension as extension was present in input")

        # Otherwise if the .xlsx file does not end with the .xlsx extension
        else:
            # Add the timestamp and .xlsx extension
            self.xlfilefull = self.xlfile + self.currentdate.strftime(" %d-%m-%Y") + ".xlsx"
            # DEBUG
            logging.debug("Added timestamp and .xlsx extension as extension was not present in input")
        logging.info("End of extensioncheck()")

        return True

    # Converter func.
    def converter(self):
        logging.info("Start of converter()")
        try:
            # Set workbook to path of .csv file,. Makes numbers in .csv input file as number fields in output excel file
            workbook = xlsxwriter.Workbook(self.mypath + str(self.xlfilefull), {'strings_to_numbers': True})

            # DEBUG
            logging.debug("Workbook set up, path, name and current date/time set without error")

        except Exception as exc:
            # DEBUG
            logging.critical("Could not set up workbook, the following error occured: %s" % exc)
            # USER OUTPUT
            print("Could not set up workbook, the following error occured:")
            print(exc)
            return False

        try:
            # Sets the worksheet in the workbook to the same name as the new excel file
            worksheet = workbook.add_worksheet(self.xlfile)
            # DEBUG
            logging.debug("Worksheet set up with name " + str(self.xlfile))

        except Exception as exc:
            # DEBUG
            logging.critical("Could not set up worksheet, the following error occured: %s" % exc)
            # USER OUTPUT
            print("Could not set up worksheet, the following error occured:")
            print(exc)
            return False

        # Sets the starting row and column to 0 so .xlsx includes all .csv data
        row = 0
        col = 0
        # DEBUG
        logging.debug("Row and column set (0)")

        # Enters the try/except, first attempts to convert the file and print a finished message
        try:
            # Whilst the file is open as read ('r') as file
            with open(self.csvfile, 'r') as file:

                # For every row 'line' in .csv file 'file'
                for line in file:

                    # Split the entries in the line by the user set delimiter 'csvdelim'
                    entry = line.split(self.csvdelim)

                    # And write to the worksheet
                    worksheet.write_row(row, col, entry)

                    # Add +1 to the row variable to loop to the next line
                    row += 1

                # DEBUG
                logging.debug("Finished writing to .xlsx document")
                # Once finished, close the workbook
                workbook.close()
                # DEBUG
                logging.debug("Workbook closed")

        except IOError as exc:
            # DEBUG
            logging.critical("The .csv file could not be found: %s" % exc)
            # USER OUTPUT
            print("The .csv file could not be found. Did you type the extensions (.csv)?")
            print(exc)
            return False

        except KeyboardInterrupt as exc:
            # DEBUG
            logging.critical("Operation interrupted by user: %s" % exc)
            # USER OUTPUT
            print("Operation interrupted by user.")
            print(exc)
            return False

        except Exception as exc:
            # DEBUG
            logging.critical("Could not perform conversion, the following error occured: %s" % exc)
            # USER OUTPUT
            print("Could not perform conversion, the following error occured:")
            print(exc)
            return False
        logging.info("End of extensioncheck()")

        return True


    def openlogger(self):
        # To add date to log file .txt
        logdate = datetime.datetime.now()

        # Sets logging to include debug functionality
        logging.basicConfig(filename="debug " + logdate.strftime("%d-%m-%y %H-%M-%S") + ".txt", level=logging.DEBUG,
                            format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")

        # DEBUG
        logging.info("Logger started")

        return True

    def closelogger(self):
        # DEBUG
        logging.info("End of closelogger()")
        pass

    def runall(self):
        # Begin logger function
        self.openlogger()
        logging.info("openlogger() finished")
        logging.info("Calling extensioncheck()")
        # Check and correct extensions
        self.extensioncheck()
        logging.info("extensioncheck() finished")
        # Run conversion function
        logging.info("Calling converter()")
        self.converter()
        logging.info("Converter() finished")
        # Close logger function
        logging.info("Calling closelogger()")
        self.closelogger()

        return True


def main():

    # Create argparse parser and set description
    parser = argparse.ArgumentParser(description="Convert a .cvs file to .xlsx")
    # Add parser arguments, uses [-f USERIN] format
    parser.add_argument('-i', '--input', type=str, help="File to convert (incl. .csv)", required=True)
    parser.add_argument('-d', '--delimeter', type=str, help="Delimeter used in .csv file (usually ,)", required=True)
    parser.add_argument('-o', '--output', type=str, help="Name for converted .xlsx file", required=True)

    # Sets argparse args
    args = parser.parse_args()

    # Sets c2e variable to CSVConverter function using the command line set args
    c2e = CSVConverter(args.input, args.delimeter, args.output)

    # Passes c2e variable to runall() function with the command line arguments
    c2e.runall()

    return True


# Main
if __name__ == '__main__':
    main()
