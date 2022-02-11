MoneyCounter 2.5 - ReadMe file

1. About

MoneyCounter is simple application written in Python to prepare monthly expenses summary on the base of export csv files from polish PKO BP and MBANK netbanks. 
Each expense is categorized - on the basis of dictionary of payment points and payment titles. 
If expense is not assigned automatically, application is asking user to define category manually. 
User can also add category to dictionary so next time, this expense will be categorized automatically. 
Summary of expenses is generated in xlsx file named when application is starting.

2. Usage

Application uses dictionary - "dict.csv" to store payment points and payment titles and categories to be automatically assigned. You can edit this file manually.
Application uses dictionary of categories - "categories.csv" to store categories to which each expense can be assigned. There must be last category called "income" which is calculated differently (it is not expense but income)
Application is able to process 3 files - 2 csv from MBANK and 1 csv from PKO BP. 
Application has two  files attached: test_mbank.csv and test_pko.csv. 

2.1 Test scenarios

0. Add 3 random categories to "categories.csv" file
1. Run controller.py with test_mbank.csv twice and test_pko.csv as test files
2. Assign categories of choice
3. For one of payments select Y - yest when asked if to save cateogry to dictionary
4. File summarized
5. Check summary file - are all categories in? are everything properly calulated
6. check dict.csv file - is there cateogry you added?

3. Plans for development

- add option to print all results to log window
- add one letter keyboard shortcuts for dropdown with categories

4. Author

MoneyCounter was developed by JPP- contact: jpp@int.pl

5. Versions

- 2.5 - make list of categories configurable in external file, all app translatd into english. Refactor of all components to be ready for publishing
- 2.4 - fixed bug with numbers in payment title - added conversion to string.
- 2.3 - fixed saving entered category for pkobp files-  it should in column M not H, make erorr messages more meaningful for end user
- 2.2 - removed graphic interface, only CLI, fixed bug with PKO categories, fixed issue with card costs in PKO files
- 2.1 - error handling added and some comments, also popups made meaninigful
- 2.0 - first release
