import requests
import xlsxwriter


class Main:
    def __init__(self, _url):
        self.url = url
        self.states_list = ["JH", "KD", "KT", "MK", "NS", "PH", "PG", "PK", "PR", "SG", "TG", "WPKL", "WL", "WPP"]
        self.state_result = []
        self.results = {}
        self.codes = []



    @staticmethod
    def get_response(url, payload):
        counter = 0
        while counter < 5:
            try:
                headers = {'Content-Type': 'application/x-www-form-urlencoded'}
                response = requests.post(url, data=payload, headers=headers, timeout=120).json()
                return response
            except:
                counter += 1
        else:
            raise Exception("Website did not load properly")



    def get_pages(self, state):
        payload = {
            "name": "",
            "alphabet": "",
            "searchtype[0]": "Firm",
            "state": state,
            "city": "",
            "keyword": "",
        }
        response = self.get_response(self.url, payload)
        return response["data"]["firms"]["numpage"]



    def run(self):
        for state in self.states_list:
            counter = 0
            current_page = 0
            pages = self.get_pages(state)
            self.state_result = []
            while counter < pages:
                current_page += 1
                print(f"[{state}] Current page {current_page} of {pages}")
                payload = {
                    "name": "",
                    "alphabet": "",
                    "searchtype[0]": "Firm",
                    "state": state,
                    "city": "",
                    "keyword": "",
                    "page": f"{current_page}"
                }
                response = self.get_response(self.url, payload)
                self.parse_data(response)
                counter += 1
        self.write_to_xlsx(self.results)



    def parse_data(self, response):
        law_firms = response["data"]["firms"]["data"]
        for law_firm in law_firms:
            code = law_firm["code"]
            city = law_firm["city"]
            state = law_firm["state"]
            contact_no = law_firm["tel1"]
            name = law_firm["name"]
            address = "{} {} {}, {} {}, {}".format(
                    law_firm["add1"], law_firm["add2"], law_firm["add3"], law_firm["postcode"], city, state)
            address = " ".join(address.split())
            fax = law_firm["fax"]
            email = law_firm["email"]
            members = []
            for member in law_firm["lawyerlist"]:
                person_name = member["name"]
                members.append(person_name)
            members = ", ".join(members)
            web = ""
            last_update = ""
            if code in self.codes:
                continue
            self.codes.append(code)
            if state not in self.results:
                self.results[state] = [[name, city, state, contact_no, name, city.upper(), address.upper(),
                                        contact_no, fax, email, web, members, last_update]]
            else:
                self.results[state].append([name, city, state, contact_no, name, city.upper(), address.upper(),
                                            contact_no, fax, email, web, members, last_update])



    @staticmethod
    def write_to_xlsx(to_xlsx):
        with xlsxwriter.Workbook('final.xlsx') as workbook:
            for key, value in to_xlsx.items():
                worksheet = workbook.add_worksheet()
                for row_num, data in enumerate(value):
                    worksheet.write_row(row_num, 0, data)


if __name__ == '__main__':
    url = "https://legaldirectory.malaysianbar.org.my/a/Lawyer/getSearchResult"
    main = Main(url)
    main.run()
