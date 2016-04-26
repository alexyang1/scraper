import sys, getopt
from linkedin_scraper import parse_linkedin
import datetime, time
import xlwt
import schedule as sched
import re


class Driver:

    def __init__(self):
        self.company_list = []
        self.filter_words = []
        self.get_company_list()
        self.get_filter_words()

        self.no_results_companies = []

        self.all_results = []
        self.approved_results = []
        self.filtered_results = []

        self.prev_approved_results = []
        self.new_results = []

    def get_company_list(self):
        self.company_list.clear()

        company_file = open('companies.txt', 'r')
        for line in company_file.readlines():
            self.company_list.append(line.strip())
        company_file.close()

        self.filter_company_names()

    def get_filter_words(self):
        self.filter_words.clear()

        filter_file = open('filter_words.txt', 'r')
        for line in filter_file.readlines():
            self.filter_words.append(line.strip())
        filter_file.close()

        self.filter_words = list(filter(None, self.filter_words))

    def filter_company_names(self):
        new_company_list = []
        for company in self.company_list:
            new_company = []

            in_parentheses = False
            for part in company.split():
                if in_parentheses and ')' not in part:
                    continue
                if ')' in part:
                    in_parentheses = False
                    continue
                if part.lower() not in ['co.', 'co', 'mgmt', 'corp.', 'corp', '&']:
                    if '(' in part:
                        in_parentheses = True
                        continue
                    new_company.append(part.strip(','))

            new_company_list.append(' '.join(new_company))

        self.company_list = list(filter(None, new_company_list))

    def new_run(self):
        self.no_results_companies.clear()

        self.prev_approved_results = self.approved_results

        self.all_results.clear()
        self.approved_results.clear()
        self.filtered_results.clear()

        for company in self.company_list:
            search_results = parse_linkedin(company, 50, 3)
            self.filter_all_results(search_results, company)
            self.all_results += search_results
            if not search_results:
                self.no_results_companies.append(company)

        self.store_results_in_file()

        print('\nTotal result(s) before filter:\t', len(self.all_results))
        print('Total result(s) filtered:\t', len(self.filtered_results))
        print('Total result(s) after filter:\t', len(self.approved_results))

    def find_new(self):
        self.new_results.clear()
        if len(self.prev_approved_results) == 0:
            self.read_prev_results_from_file()

        for result in self.approved_results:
            is_new = True
            for prev_result in self.prev_approved_results:
                if result['Title'] == prev_result['Title'] and result['Company'] == prev_result['Company']:
                    is_new = False

            if is_new:
                self.new_results.append(result)

        print('Total new result(s):\t\t', len(self.new_results))

    def filter_all_results(self, results, company):
        for result in results:
            filter_booleans = self.filter_result(company, result)

            if filter_booleans == (True, True):
                self.approved_results.append(result)
            elif filter_booleans == (True, False):
                result['Reason'] = 'Failed company filter - \"' + company + '\"'
                self.filtered_results.append(result)
            elif filter_booleans == (False, True):
                result['Reason'] = 'Failed title filter'
                self.filtered_results.append(result)
            else:
                result['Reason'] = 'Failed both filters'
                self.filtered_results.append(result)

    def filter_result(self, company, x):
        right_title = True
        title_parts = x['Title'].split()
        for word in self.filter_words:
            if ' ' in word:
                if word.lower() in x['Title'].lower():
                    right_title = False
                    break
            else:
                for part in title_parts:
                    if word.lower() == part.lower():
                        right_title = False
                        break

        right_company = False
        num_match = 0
        max_match = min(len(company.split()), 2, len(x['Company'].split()))
        for part in company.split():
            if part.lower() in ['and', 'co.', 'co', 'mgmt', 'corp.', 'corp', '&']:
                continue
            if part.lower() in x['Company'].lower():
                num_match += 1
                if num_match >= max_match:
                    right_company = True
                    break

        return right_title, right_company

    def remove_no_results(self):
        for company in self.company_list:
            if company in self.no_results_companies:
                self.company_list.remove(company)

    def print_to_excel(self, filename):
        book = xlwt.Workbook(encoding="utf-8")
        sheet1 = book.add_sheet("Results")

        sheet1.write(0, 0, "Job Title")
        sheet1.write(0, 1, "Company")
        sheet1.write(0, 2, "Posted")
        sheet1.write(0, 3, "URL")

        i = 1
        for result in self.approved_results:
            sheet1.write(i, 0, result['Title'])
            sheet1.write(i, 1, result['Company'])
            sheet1.write(i, 2, result['Date'])
            sheet1.write(i, 3, result['Link'])
            i += 1

        sheet2 = book.add_sheet("Filtered")

        sheet2.write(0, 0, "Job Title")
        sheet2.write(0, 1, "Company")
        sheet2.write(0, 2, "Posted")
        sheet2.write(0, 3, "URL")
        sheet2.write(0, 4, "Reason")

        i = 1
        for result in self.filtered_results:
            sheet2.write(i, 0, result['Title'])
            sheet2.write(i, 1, result['Company'])
            sheet2.write(i, 2, result['Date'])
            sheet2.write(i, 3, result['Link'])
            sheet2.write(i, 4, result['Reason'])
            i += 1

        sheet3 = book.add_sheet("New Results")

        sheet3.write(0, 0, "Job Title")
        sheet3.write(0, 1, "Company")
        sheet3.write(0, 2, "Posted")
        sheet3.write(0, 3, "URL")

        i = 1
        for result in self.new_results:
            sheet3.write(i, 0, result['Title'])
            sheet3.write(i, 1, result['Company'])
            sheet3.write(i, 2, result['Date'])
            sheet3.write(i, 3, result['Link'])
            i += 1

        book.save(filename)

    def store_results_in_file(self):
        outfile = open('previous_results.txt', 'w')
        i = 0
        for result in self.approved_results:
            outfile.write(str(i) + '\n')
            outfile.write(result['Title'] + '\n')
            outfile.write(result['Company'] + '\n')
            outfile.write(str(result['Date']) + '\n')
            outfile.write(result['Link'] + '\n')
            i += 1

        outfile.close()

    def read_prev_results_from_file(self):
        infile = open('previous_results.txt', 'r')
        while infile.readline() is not '':
            result = dict()
            result['Title'] = infile.readline().strip()
            result['Company'] = infile.readline().strip()
            result['Date'] = infile.readline().strip()
            result['Link'] = infile.readline().strip()
            self.prev_approved_results.append(result)

        infile.close()


def main(argv):

    def usage():
        print('usage: python3 driver.py [options]')
        print('\toptions:')
        print('\t   -h, --help\t: Displays usage')
        print('\t   -s\t\t: Runs driver on a schedule')
        print('\t   -o, --output\t: Output excel file name (Optional)')
        print('\t   -d\t\t: Debug')

    date = datetime.date.today().strftime("%d %B %Y")
    filename = 'LinkedIn - ' + date + '.xls'

    use_schedule = False

    def parse():
        driver = Driver()
        driver.new_run()
        driver.find_new()
        driver.print_to_excel(filename)

    def schedule():
        sched.every().monday.at('5:00').do(parse)

        while True:
            sched.run_pending()
            print('\n' + str(sched.jobs[0]))

            next_run_nums = re.findall('\d+', str(sched.jobs[0]).split('next run:')[1])
            future = datetime.datetime(int(next_run_nums[0]), int(next_run_nums[1]), int(next_run_nums[2]),
                                       int(next_run_nums[3]), int(next_run_nums[4]), int(next_run_nums[5]))

            today = datetime.datetime.today()
            while today < future:
                today = datetime.datetime.today()
                time.sleep(1)

    try:
        opts, args = getopt.getopt(argv, "hso:d", ['help, output='])
    except getopt.GetoptError:
        usage()
        sys.exit(2)

    for opt, arg in opts:
        if opt in ('-h', '--help'):
            usage()
            sys.exit()
        if opt == '-s':
            use_schedule = True
        if opt in ('-o', '--output'):
            filename = arg
        if opt == '-d':
            global _debug
            _debug = 1

    if use_schedule:
        schedule()
    else:
        parse()


if __name__ == '__main__':
    main(sys.argv[1:])
