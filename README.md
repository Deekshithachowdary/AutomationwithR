Overview
This repository contains R scripts designed to automate the process of generating daily backorder data reports. The automation significantly reduces manual effort, saving over 2 hours of work each day. The scripts handle large volumes of data, processing more than 700,000 lines of Excel data, with the data volume increasing each week.

Features

Automated Data Processing: Efficiently handles large datasets in Excel, ensuring accurate and timely data processing.
Report Generation: Automatically generates comprehensive backorder reports, reducing manual intervention.
Time-saving: Saves over 2 hours of manual work each day, increasing productivity.
Scalable: Capable of handling increasing data volumes, ensuring the solution remains effective as data grows.

Install Required Packages:
Make sure you have R installed on your system. Then, install the necessary packages by running:

nstall Required Packages:
Make sure you have R installed on your system. Then, install the necessary packages by running:


install.packages("readxl")
install.packages("openxlsx")
install.packages("dplyr")
install.packages("lubridate")
install.packages("tidyverse")
install.packages("tidyr")
install.packages("pivottabler")
install.packages("rpivotTable")
install.packages("magrittr")
Load Required Libraries:
Add the following lines to your script to load the necessary libraries:


library("readxl")
library("openxlsx")
library("dplyr")
library("lubridate")
library("tidyverse")
library("tidyr")
library("pivottabler")
library("rpivotTable")
library("magrittr")


Prepare Your Data: Ensure your Excel data files are properly formatted and placed in the data directory.
Project Structure
data/: Directory containing input Excel files.
scripts/: Contains the main R scripts for data processing and report generation.
output/: Directory where the generated reports are saved.
README.md: This file.

Acknowledgements
This project utilizes the readxl, openxlsx, dplyr, lubridate, tidyverse, tidyr, pivottabler, rpivotTable, and magrittr packages for data processing and report generation.
