from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from time import sleep
from re import findall
import xlwt

base_url = 'http://www.linkedin.com/jobs/'


def parse_linkedin(company_name=None, max_results=50, number_of_retrials=0):

    browser = webdriver.PhantomJS(executable_path='./phantomjs')
    browser.get(base_url)
    print('\nBeginning LinkedIn.com search for', company_name, '...')

    try:
        search_box = browser.find_element_by_class_name('job-search-field')
    except NoSuchElementException:
        browser.close()
        print('Connection failure')
        if number_of_retrials > 0:
            print('Retrying search ...')
            sleep(1)
            return parse_linkedin(company_name, max_results, number_of_retrials - 1)
        else:
            return []

    search_box.send_keys(company_name)
    search_box.send_keys(Keys.RETURN)

    print(browser.current_url)

    try:
        search_count = browser.find_element_by_xpath('.//div[@class = "results-context"]').text
    except NoSuchElementException:
        try:
            browser.find_element_by_xpath('.//div[@class = "jserp-page-results empty"]')
            print('Search query returned no results')
            browser.close()
            return []
        except NoSuchElementException:
            browser.close()
            print('Connection failure')
            if number_of_retrials > 0:
                print('Retrying search...')
                return parse_linkedin(company_name, max_results, number_of_retrials - 1)
            else:
                return []

    job_numbers = findall('\d+', search_count)

    if len(job_numbers) > 1:    # If number of results > 1000, merge numbers separated by comma
        num_results = 1000 * int(job_numbers[0]) + int(job_numbers[1])
    else:                       # Otherwise, number of results will appear in job_numbers[0]
        num_results = int(job_numbers[0])

    print('Search query for', company_name, 'returned', num_results, 'results')

    # Calculate number of pages
    if max_results is not None:     # If upper bound of results is given
        num_results = min(max_results, num_results)

        num_pages = int(num_results / 25)
        if (num_results % 25) is not 0:
            num_pages += 1
        print('Parsing', num_results, 'results across', num_pages, 'pages')
    else:                           # Otherwise just use number of results returned
        num_pages = int(num_results / 25)
        if (num_results % 25) is not 0:
            num_pages += 1
        print('Parsing', num_results, 'results across', num_pages, 'pages')sc

    listings = []   # Final list of dictionary objects, to be returned at end of function
    for i in range(num_pages):

        # Build URL in iterable form
        start_num = 25 * i
        company_url = company_name.split()
        company_url = '+'.join(word for word in company_url if word is not '&')

        url = ''.join([base_url, 'search?keywords=', company_url, '&start=', str(start_num), '&count=25'])
        print(url)
        browser.get(url)
        sleep(1)

        results = browser.find_elements_by_class_name('job-listing')

        if (start_num + 25) > num_results:  # If we want <25 results on the page (last page)...
            num = num_results - start_num   # ... calculate number of results we want on the page
            try:
                assert(num <= len(results))
            except AssertionError:
                num = len(results)

            print('Parsing', num, 'results from page', i+1)
            for j in range(num):
                listing = dict()

                listing['Link'] = results[j].\
                    find_element_by_xpath('.//a[@class = "job-title-link"]').get_attribute('href')
                listing['Title'] = results[j].\
                    find_element_by_xpath('.//span[@class = "job-title-text"]').text
                listing['Company'] = results[j].\
                    find_element_by_xpath('.//span[@class = "company-name-text"]').text
                listing['Location'] = results[j].\
                    find_element_by_xpath('.//span[@itemprop = "addressLocality"]').text
                try:
                    listing['Date'] = int(findall('\d+', results[j].
                                                  find_element_by_xpath('.//span[@class = "job-date-posted"]').text)[0])
                except NoSuchElementException:
                    listing['Date'] = 0

                listings.append(listing)
            continue

        print('Parsing', len(results), 'results from page', i+1)
        for result in results:
            listing = dict()

            listing['Link'] = result.find_element_by_xpath('.//a[@class = "job-title-link"]').get_attribute('href')
            listing['Title'] = result.find_element_by_xpath('.//span[@class = "job-title-text"]').text
            listing['Company'] = result.find_element_by_xpath('.//span[@class = "company-name-text"]').text
            listing['Location'] = result.find_element_by_xpath('.//span[@itemprop = "addressLocality"]').text
            try:
                listing['Date'] = int(findall('\d+', result.
                                              find_element_by_xpath('.//span[@class = "job-date-posted"]').text)[0])
            except NoSuchElementException:
                listing['Date'] = 0

            listings.append(listing)

        wait_time = 1
        sleep(wait_time)

    browser.close()
    try:
        assert(len(listings) == num_results)
    except AssertionError:
        print('Number of parsed results is not equal to number of results returned by search')
        if number_of_retrials > 0:
            print('Retrying search...')
            return parse_linkedin(company_name, max_results, number_of_retrials - 1)
        else:
            return []

    return listings


def filter_result(company, x):

    filter_file = open('filter_words.txt', 'r')
    filter_words = []
    for line in filter_file.readlines():
        filter_words.append(line.strip())
    filter_words = list(filter(None, filter_words))
    filter_file.close()

    right_title = True
    title_parts = x['Title'].split()
    for word in filter_words:
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
        if part.lower() is 'and' or part.lower() is '&' or part.lower() in 'co.':
            continue
        if part.lower() in x['Company'].lower():
            num_match += 1
            if num_match >= max_match:
                right_company = True
                break

    return right_title, right_company


def print_to_excel(approved, filtered):
        # Export results to Excel
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Results")

    sheet1.write(0, 0, "Job Title")
    sheet1.write(0, 1, "Company")
    sheet1.write(0, 2, "Location")
    sheet1.write(0, 3, "Posted")
    sheet1.write(0, 4, "URL")

    i = 1
    for result in approved:
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
    sheet2.write(0, 5, "Reason")

    i = 1
    for result in filtered:
        sheet2.write(i, 0, result['Title'])
        sheet2.write(i, 1, result['Company'])
        sheet2.write(i, 2, result['Location'])
        sheet2.write(i, 3, result['Date'])
        sheet2.write(i, 4, result['Link'])
        sheet2.write(i, 5, result['Reason'])
        i += 1

    book.save('LinkedIn.xls')


def main():

    company_file = open('companies.txt', 'r')
    company_list = []
    for line in company_file.readlines():
        company_list.append(line.strip())
    company_list = list(filter(None, company_list))
    company_file.close()

    all_results = []
    approved_results = []
    filtered_results = []

    for company in company_list:
        trial_results = parse_linkedin(company_name=company, number_of_retrials=3)
        all_results += trial_results

        if len(trial_results) is 0:
            continue

        print('Filtering results for', company, '...')
        for result in trial_results:

            filter_status = filter_result(company, result)

            if filter_status == (True, True):
                approved_results.append(result)
            elif filter_status == (False, True):
                result['Reason'] = 'Failed title filter'
                filtered_results.append(result)
            elif filter_status == (True, False):
                result['Reason'] = 'Failed company filter'
                filtered_results.append(result)
            else:
                result['Reason'] = 'Failed both filters'
                filtered_results.append(result)

    print('\nTotal result(s) before filter:\t', len(all_results))
    print('Total result(s) filtered:\t', len(filtered_results))
    print('Total result(s) after filter:\t', len(approved_results))

    print_to_excel(approved_results, filtered_results)

if __name__ == "__main__":
    main()
