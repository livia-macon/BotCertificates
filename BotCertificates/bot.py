from botcity.core import DesktopBot
from botcity.plugins.excel import BotExcelPlugin
from collections import namedtuple
import datetime
import numpy as np

class Bot(DesktopBot):
    def action(self, execution=None):
        # Instantiate the plugin an reading sheet
        sheet = BotExcelPlugin().read(self.get_resource_abspath("certificates_list.xlsx"))

        # Replacing cell nan for empty
        sheet._sheets[sheet.active_sheet].replace(np.nan, '', inplace=True)

        # Returns the contents of an entire sheet in a list of lists format.
        data = sheet.as_list()[1:]

        # Declaring namedtuple() == DataForCerificate(Name='', Email='',Process='', Date='', Index='')
        DataForCertificate = namedtuple('DataForCerificate', ['Name','Email', 'Process', "Date", "Index"] )
        data_certificate = [DataForCertificate(data[0], data[1], data[2], data[3], index)  for index, data in enumerate(data, start=2)]

        for datas in data_certificate:
            # If process field is empty
            if not datas.Process:
                # Open Power Point with certificate model
                self.execute(self.get_resource_abspath("certificate_model.pptx"))
                # verify if windown open
                if not self.find( "bot", matching=0.97, waiting_time=15000):
                    self.not_found("bot")
                # Find text "cetify"
                if not self.find( "certify", matching=0.97, waiting_time=15000):
                    self.not_found("certify")
                # Click relative to previously found image
                self.click_relative(42, 73)
                # Type the name in field
                self.kb_type(datas.Name)
                # Find text "certification date"
                if not self.find( "certification_date", matching=0.97, waiting_time=10000):
                    self.not_found("certification_date")
                # Click relative to previously found image
                self.click_relative(97, -50)
                # Type the date in field
                self.kb_type(datetime.datetime.now().strftime("%d/%m/%Y"))
                # Press shortcut to go bar of actions
                self.type_keys(["alt", "q"])
                # Type the text "save as"
                self.kb_type("save as")
                self.enter()
                # Type the name to file
                self.kb_type(f'{str(datas.Name).replace(" ", "_")}_{datetime.datetime.now().strftime("%d_%m_%Y")}')
                self.tab()
                # Choice the file format to save
                # Type "p" and after press key down (arrow down) untill arrive pdf format
                self.type_key("p")
                self.type_down()
                self.type_down()
                self.type_down()
                self.enter()
                self.enter()
                # If find an error message, take a screenshot and write in the worksheet "Error generating certificate".
                if self.find( "save_as_error", matching=0.97, waiting_time=2000):
                    self.save_screenshot(str(datas.Name).replace(" ", "_") + '_' + datetime.datetime.now().strftime("%d_%m_%Y") + "_.png")
                    sheet.set_cell("C", datas.Index, "Error when generating certificate")
                    print(f"It was not possible generate the certificate of {datas.Name}")
                    # closes windown press key rigth to dont save changed
                    self.alt_f4()
                    self.alt_f4(wait=2000)
                    self.alt_f4()
                    self.type_right()
                    self.enter()
                    continue
                # If was possible saves, close windown and press key rigth to dont save changed
                self.alt_f4()
                self.type_right()
                self.enter()
                # Edit worksheet
                sheet.set_cell("C", datas.Index, "Processed")
                sheet.set_cell("D", datas.Index, datetime.datetime.now().strftime("%d/%m/%Y"))  
                sheet.write(self.get_resource_abspath("certificates_list.xlsx"))
                print(f"Certificate of {datas.Name} emited today.")
            else:
                print(f"Certificate of {datas.Name} has already been emited in {datas.Date}")
    
    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()
