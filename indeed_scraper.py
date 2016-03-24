from bs4 import BeautifulSoup
from urllib import request
import re
from time import sleep
import xlwt

base_url = 'http://www.indeed.com'


def parse_indeed(company_name=None, max_results=50, number_of_retrials=0):

    print('\nBeginning Indeed.com search for', company_name, '...')

    # Build URL for initial search page
    if company_name is not None:
        company_name_url = company_name.split()
        company_name_url = '+'.join(word for word in company_name_url)
    else:
        return []

    url = ['http://www.indeed.com/jobs?q=', company_name_url, '&l=']
    final_url = ''.join(url)
    print(final_url)

    # Open URL and read HTML into BeautifulSoup object
    try:
        html_encoded = request.urlopen(final_url).read()
    except request.URLError:
        print('Connection Failure')
        if number_of_retrials > 0:
            print('Retrying search...')
            return parse_indeed(company_name, max_results, number_of_retrials - 1)
        else:
            return []

    soup = BeautifulSoup(html_encoded, 'lxml')

    if soup.find(id='searchCount') is None:
        print('Search query returned no results')
        if number_of_retrials > 0:
            print('Retrying search...')
            return parse_indeed(company_name, max_results, number_of_retrials - 1)
        else:
            return []

    # Read number of job listings returned from 'searchCount' object in HTML
    search_count = soup.find(id='searchCount').string
    job_numbers = re.findall('\d+', search_count)

    if len(job_numbers) > 3:
        num_results = (int(job_numbers[2]) * 1000) + int(job_numbers[3])
    else:
        num_results = int(job_numbers[2])

    print('Search query returned', num_results, 'results')

    # Calculate number of pages from number of job listings - 10 per page
    if max_results is not None:     # If upper bound of results is set
        num_results = min(max_results, num_results)  # Use min of upper bound or results returned
        num_pages = int(num_results / 10)
        if (num_results % 10) is not 0:
            num_pages += 1
    else:                           # Otherwise just use number of results returned
        num_pages = int(num_results / 10)
        if (num_results % 10) is not 0:
            num_pages += 1

    print('Parsing', num_results, 'results across', num_pages, 'pages')

    results = []
    for i in range(0, num_pages):
        # Build URL for each page of results
        start_num = i * 10
        current_url = ''.join([final_url, '&start=', str(start_num)])
        print(current_url)

        results += read_page(current_page=current_url, num_to_read=None, read_last_row=False)
        results += read_page(current_page=current_url, num_to_read=None, read_last_row=True)

        print('Parsing results from page', i+1)

        '''
        if start_num + 10 > num_results:
            num_last_page_results = num_results - start_num
            print('Parsing', num_last_page_results, 'results from page', i+1)

            page_results = read_page(current_url, None, False)
            page_results += read_page(current_url, None, True)

            print(page_results)


            page_results = read_page(current_page=current_url, num_to_read=None, read_last_row=False)
            print(page_results)
            if len(page_results) > num_last_page_results:
                results += page_results[:num_last_page_results]
            elif len(page_results) == num_last_page_results-1:
                last_row_results = read_page(current_page=current_url, num_to_read=None, read_last_row=True)
                results += page_results + last_row_results
            else:
                results += page_results

        else:
            print('Parsing 10 results from page', i+1)
            page_results = read_page(current_page=current_url, num_to_read=None, read_last_row=False)
            page_results += read_page(current_page=current_url, num_to_read=None, read_last_row=True)
            results += page_results
        '''

        wait_time = 1
        sleep(wait_time)

    results = results[:num_results]

    try:
        assert(len(results) == num_results)
    except AssertionError:
        print('Number of parsed results is not equal to number of results returned by search')
        if number_of_retrials > 0:
            print('Retrying search...')
            return parse_indeed(company_name, max_results, number_of_retrials - 1)
        else:
            return []

    return results


def read_page(current_page=None, num_to_read=None, read_last_row=False):

    # Read HTML into BeautifulSoup object
    html_encoded = request.urlopen(current_page).read()
    soup = BeautifulSoup(html_encoded, 'lxml')
    results_area = soup.find(id='resultsCol')
    if read_last_row:
        results_list = results_area.find_all(class_='lastRow row result')
    else:
        results_list = results_area.find_all(class_=' row result')

    if num_to_read is None:
        num_to_read = len(results_list)

    results = []
    for j in range(num_to_read):
        result_tag = results_list[j]
        result = dict()

        result['Title'] = result_tag.find('a', itemprop='title')['title']

        parent_company_tag = result_tag.find('span', class_='company')
        if parent_company_tag is None:
            continue

        company_tag = parent_company_tag.find('span', itemprop='name')

        if company_tag.find('a') is not None:
            if company_tag.a.find('b') is not None:
                temp = []
                for tag in company_tag.a.contents:
                    temp.append(tag.string)

                company = ''.join(temp)
                company = company.split()
                company = ' '.join(company)
            else:
                company = company_tag.a.string
        else:
            if company_tag.find('b') is not None:
                temp = []
                for tag in company_tag.contents:
                    temp.append(tag.string)

                company = ''.join(temp)
                company = company.split()
                company = ' '.join(company)
            else:
                company = company_tag.string

        result['Company'] = company.strip()

        link = base_url + result_tag.h2.a.get('href')
        result['Link'] = link

        result['Location'] = result_tag.find('span', itemprop='addressLocality').string
        result['Date'] = result_tag.find('span', class_='date').string

        results.append(result)

    return results


def main():

    company_file = open('companies.txt', 'r')
    company_list = []
    for line in company_file.readlines():
        company_list.append(line.strip())
    company_list = list(filter(None, company_list))
    company_file.close()

    filter_file = open('filter_words.txt', 'r')
    filter_words = []
    for line in filter_file.readlines():
        filter_words.append(line.strip())
    filter_words = list(filter(None, filter_words))
    filter_file.close()

    all_results = []
    approved_results = []
    filtered_results = []

    for company in company_list:
        trial_results = parse_indeed(company_name=company)
        all_results += trial_results
        print('Filtering results for', company, '...')

        for result in trial_results:

            right_title = True
            title_parts = result['Title'].split()
            for word in filter_words:
                if ' ' in word:
                    if word in result['Title'].lower():
                        right_title = False
                        break
                else:
                    for part in title_parts:
                        if word.lower() == part.lower():
                            right_title = False
                            break

            right_company = False
            num_match = 0
            max_match = min(len(company.split()), 2)
            for part in company.split():
                if part.lower() is 'and' or part.lower() is '&' or part.lower() is 'the':
                    continue

                if part.lower() in result['Company'].lower():
                    num_match += 1
                    if num_match >= max_match:
                        right_company = True
                        break

            if right_title and right_company:
                approved_results.append(result)
            else:
                filtered_results.append(result)

    print('\nTotal result(s) before filter:\t', len(all_results))
    print('Total result(s) filtered:\t', len(filtered_results))
    print('Total result(s) after filter:\t', len(approved_results))

    # Export results to Excel
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Results")

    sheet1.write(0, 0, "Job Title")
    sheet1.write(0, 1, "Company")
    sheet1.write(0, 2, "Location")
    sheet1.write(0, 3, "Posted")
    sheet1.write(0, 4, "URL")

    i = 1
    for result in approved_results:
        sheet1.write(i, 0, result['Title'])
        sheet1.write(i, 1, result['Company'])
        sheet1.write(i, 2, result['Location'])
        sheet1.write(i, 3, result['Date'])
        sheet1.write(i, 4, result['Link'])
        i += 1

    sheet2 = book.add_sheet("Filtered")

    sheet2.write(0, 0, "Job Title")
    sheet2.write(0, 1, "Company")
    sheet2.write(0, 2, "Location")
    sheet2.write(0, 3, "Posted")
    sheet2.write(0, 4, "URL")

    i = 1
    for result in filtered_results:
        sheet2.write(i, 0, result['Title'])
        sheet2.write(i, 1, result['Company'])
        sheet2.write(i, 2, result['Location'])
        sheet2.write(i, 3, result['Date'])
        sheet2.write(i, 4, result['Link'])
        i += 1

    book.save('Indeed.xls')

main()
