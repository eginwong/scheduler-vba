# Scheduler

## Objective
The scheduler spreadsheet enables users to easily schedule activities with particularly roles and capabilities.

## Features
- generates a schedule based on customizable roles and users
- optimized solutions using Solver/OpenSolver for large datasets
- priority function to avoid bias when scheduling
- unavailability rules to help plan for future vacations in scheduling

## Example
A volunteering group may require several roles to run a meeting.

*Roles*: `Team Lead`, `Organizer`, `Chaperone`, `Volunteer`.
Each role can require a certain number of candidates to fill the role. 
There are a pool of potential candidates for these roles, from A - F. 

The table (represented by the `CAPABILITIES` tab in the spreadsheet) below demonstrates which candidate has the skills to do which role, 
and how many of each role are required:

| Candidate | `Team Lead` (1)   | `Organizer` (1)   | `Chaperone` (2)   | `Volunteer` (3)   |
| --------- |:-----------------:|:-----------------:|:-----------------:|:-----------------:|
| A         | x                 |                   | x                 | x                 |
| B         | x                 |                   | x                 | x                 |
| C         |                   | x                 |                   | x                 |
| D         |                   |                   | x                 | x                 |
| E         |                   |                   | x                 |                   |
| F         |                   |                   | x                 |                   |

Given this sort of input, the scheduler can optimize and fill all the roles given candidate availability.

Let's say that candidates A and C are unavailable for the next volunteer group session. 
The scheduler can also take this into consideration and adjust the schedule accordingly. 

In order to avoid bias and having the same candidates perform the same roles, there is a priority function in the objective Function
of the linear program of the model to lean more towards candidates who have never performed the role before. 
This data can be seen in the `USERS` tab of the spreadsheet.

There's several more features so what are you waiting for? Give it a spin!

## Instructions
1. Open up the `.xlsm` file in `bin/`.
2. Follow instructions to set up data in order to run the scheduler.

## Third-party Libraries
- Excel Solver Add-In
- (optionally) [OpenSolver](https://opensolver.org/) for larger datasets

## Credit
Please make sure credit is given where credit is due, if you are interested in using or extending this scheduler implementation.