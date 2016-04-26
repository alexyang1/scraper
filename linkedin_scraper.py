from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from time import sleep
from re import findall

base_url = 'http://www.linkedin.com/jobs/'


def parse_linkedin(company_name=None, max_results=50, number_of_retrials=0):

    browser = webdriver.PhantomJS(executable_path='./phantomjs')
    browser.get(base_url)
    print('\nBeginning LinkedIn.com search for', company_name, '...')

    try:
        if (number_of_retrials / 2) is not 0:
            search_box = browser.find_element_by_class_name('job-search-field')
        else:
            search_box = browser.find_element_by_id('field-keyword-name')
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
    else:
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
        print('Parsing', num_results, 'results across', num_pages, 'pages')

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
                try:    # All these try/except statements are a little clumsy - rework later
                    listing['Link'] = results[j].\
                        find_element_by_xpath('.//a[@class = "job-title-link"]').get_attribute('href')
                    listing['Title'] = results[j].\
                        find_element_by_xpath('.//span[@class = "job-title-text"]').text
                    listing['Company'] = results[j].\
                        find_element_by_xpath('.//span[@class = "company-name-text"]').text
                except NoSuchElementException:
                    print('Could not find listing elements')
                    if number_of_retrials > 0:
                        number_of_retrials -= 1
                        j -= 1
                        continue
                    else:
                        continue
                try:
                    listing['Date'] = int(findall('\d+', results[j].
                                                  find_element_by_xpath('.//span[@class = "job-date-posted"]').text)[0])
                except NoSuchElementException:
                    listing['Date'] = 0

                listings.append(listing)
            continue

        print('Parsing', len(results), 'results from page', i+1)
        for j in range(len(results)):
            listing = dict()

            try:
                listing['Link'] = results[j].\
                    find_element_by_xpath('.//a[@class = "job-title-link"]').get_attribute('href')
                listing['Title'] = results[j].find_element_by_xpath('.//span[@class = "job-title-text"]').text
                listing['Company'] = results[j].find_element_by_xpath('.//span[@class = "company-name-text"]').text
            except NoSuchElementException:
                print('Could not find listing elements')
                if number_of_retrials > 0:
                    number_of_retrials -= 1
                    j -= 1
                    continue
                else:
                    continue
            try:
                listing['Date'] = int(findall('\d+', results[j].
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
