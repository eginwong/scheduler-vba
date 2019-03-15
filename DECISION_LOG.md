# Scheduler Design Log

Bug: there is a limit of 200 decision variables using Microsoft Excel's Solver add-in. 
- Two alternatives:
  - [Premium Solver](https://www.solver.com/premium-solver%C2%AE-platform), a paid for plug-in
  - [OpenSolver](https://opensolver.org/), an open-source solution
- Quick workaround fix by writing conditional code to use OpenSolver when detected
- TODO: add more graceful check
- TODO: include instructions on how to download, install, and configure OpenSolver in a hidden sheet


Feature Request: add vacation calendar
- a user should be able to add planned vacation days to be taken into consideration by the Solver
- ~~TODO:~~
  - create new tab for vacations
  - add conditional formatting for past days
  - add purge button
  - add add vacation button
  - create vacation macro
    - pull names for README list
    - include calendar picker
    - sort list after insertion
  - create purge old records macro
  - include vacation in solver macro
    - maybe manually toggle names to False and then back for simple mechanism
    - for each user, scan through vacation dates first


- Decided to start with adding functionality one at a time
- Discuss design - one hidden user template that is cloned for every new user
  - intention was to have each user with their own profile, but most of the data is referenced from elsewhere anyway
  - only unique piece was the last execution date for each role, which can be stored much more succinctly. Requires redesign.

- update README
  - good idea to write out design details or overall list of features to keep your scope in mind as you're building
- create macro for new roles
  - develop ideal output and update over time
  - write macro to automate the process, updating spreadsheets as required
- create macro to run solver
  - found an online forum talking about safe coding practices when using references, Microsoft Excel Add-Ins
  - added the check that was recommended but causes problems on startup of excel, type mismatch
    - figured out that the code snippet example was wrong and adjusted to proper VBA syntax
- create macro to update user profiles with last run date
  - originally, had multiple spreadsheets for each user profile but eventually refactored into a single users spreadsheet for aggregation
- delete roles
  - deleting is always so much easier
  - had to make a hack for this one as I did not want to scan through all user sheets so I hard-coded references for the first 20 roles
  - chose 20 because anything higher would cause a lag when cloning and resetting proper values and references for new users
  - sucky limitation but best for POC needs
  - instead of hard deleting the column with the row, instead clear it
  - discovered that that situation does not work if the role pre-existed and would then leave a gaping hole between roles, which would break functionality more
- Make excel more abstract					
- Error conditions for duplicate names					*check if user/role name already exists, and let the user know
  - search for proposed username in list, return boolean result and act accordingly
- Fix to have only one user sheet; afterwards, remove limitation on roles
  - take backup and remove pieces one-at-a-time and slowly refactor logically
- Can't clear role, what if we delete an established role, there would be a gap					will be fixed with converting to only having one user sheet
- Error conditions for special names (Priority, etc etc.), whenever I remove roles/users	 
  - went through spreadsheets trying to remove any unnecessary text in columns that would be searched
  - along the way, refactored the null check for Find methods, to keep code clean
  - removed row and broke some hard-coded code, went through to sort that out 
- Bug fixes:
  - freezing panes for better visibility
  - UX updates for better placed buttons
  - copying formatting for row/column insertion
  - bugs from cells no longer being labelled
