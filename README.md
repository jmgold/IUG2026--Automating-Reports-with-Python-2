# IUG2026: Automating Reports with Python 2
Script examples referenced within Automating Reports with Python 2.  Presented at [the IUG2026 Annual Conference](https://www.innovativeusers.org/iug_2026.php)

Scripts and queries written by Jeremy Goldstein, building upon the work of Gem Stone-Logan (with her kind permission) and her presentation [Automating Reports with Python](https://www.gemstonelogan.com/presentations.html), presented at the 2017 and 2018 IUG Annual Conferences.

## Prerequisites:
All scripts are run in an environbment labeled py313 defined by [py313.yml](https://github.com/jmgold/IUG2026--Automating-Reports-with-Python-2/blob/main/py313.yml).

Scripts also rely on local information that must filled out within [config.ini](https://github.com/jmgold/IUG2026--Automating-Reports-with-Python-2/blob/main/py313.yml), which contains placeholder information as a starting point.

## The Scripts
1. [weeklynew_2026.py](https://github.com/jmgold/IUG2026--Automating-Reports-with-Python-2/blob/main/weeklynew_2026.py) - Updated version of the final script from Gem Stone-Logan's original presentation to utilize functions and a config file.  Script uses [WeeklyNewItemsRev.sql](https://github.com/jmgold/IUG2026--Automating-Reports-with-Python-2/blob/main/WeeklyNewItemsRev.sql), which has been slightly revised from Gem's query to eliminate potions that would not yield results when searching Minuteman's database.
2. [weeklynew_2026_inline_query.py](https://github.com/jmgold/IUG2026--Automating-Reports-with-Python-2/blob/main/weeklynew_2026_inline_query.py) - Script altered to include sql query within the python code instead of reading an external file.  Query is also updated to take a library location code as a parameter passed to the run_query() function.
3. [weeklynew_2026_csv.py](https://github.com/jmgold/IUG2026--Automating-Reports-with-Python-2/blob/main/weeklynew_2026_csv.py) - replaces write_excel() function with write_csv() to provide a means of producing file outputs without the need to customize the formatting for different queries.
