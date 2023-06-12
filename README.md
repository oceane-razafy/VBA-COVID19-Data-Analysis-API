# COVID Data Analysis with VBA

A VBA project for analyzing the impact of COVID-19 across countries and comparing key metrics, developed as part of a master's degree coursework at Paris Dauphine University by another student and myself. Completed in October 2021 and received a grade of 17/20.

In order to make the VBA code visible on GitHub, it is stored in separate `.bas` files within the 'modules' directory of this repository.

## Table of Contents

- [Project Description](#project-description)
- [Key Features](#key-features)
- [Getting Started](#getting-started)
- [Acknowledgments](#acknowledgments)

## Project Description

This VBA project focuses on analyzing the impact of COVID-19 across different countries. The data used in this project is sourced from "The Johns Hopkins University" and is updated daily. It provides information on infections, deaths, and death rates for the past 31 days. The purpose of this project is to allow users to compare the COVID-19 impact on the population among three countries of their choice.

The main feature of this module is the ability to update data for each country selected by the user from the "Extraction des données" sheet. Users can view all available countries and make their selections from the dropdown lists provided.

COVID-19 has been a global topic of concern, and we chose these continuously updated data to provide a real-time understanding of the current situation. The dataset contains over 6000 records, offering users a comprehensive view of COVID-19 statistics in countries that may not receive significant media coverage. This allows users to identify the most affected regions, countries, or continents. The data is obtained in JSON format from the following link:
[https://coronavirus.politologue.com/data/coronavirus/coronacsv.aspx?format=json](https://coronavirus.politologue.com/data/coronavirus/coronacsv.aspx?format=json).

We also explored [https://api.covidtracking.com/v1/states/daily.json](https://api.covidtracking.com/v1/states/daily.json), but it focuses only on the United States, whereas we aimed to provide a global perspective.

## Key Features

- Interactive data update: Users can select countries from dropdown lists to update data for analysis.
- Comparative analysis: Users can compare the impact of COVID-19 across three selected countries.
- Real-time data: The dataset is updated daily to provide the most current information available.

## Getting Started

To use this VBA project, follow these steps:

1. Clone the repository or download the project files.
2. Open the Excel file (.xlsm) containing the VBA code and data.
3. Enable macros and content if prompted.
4. Navigate to the "Extraction des données" sheet.
5. Select three countries of interest from the dropdown lists.
6. Click the "Extraction" button to retrieve and update the COVID-19 data for the selected countries.
7. Explore the updated data and analyze the impact of COVID-19.

## Acknowledgments

The COVID-19 data used in this project is sourced from "The Johns Hopkins University" and is made available through the following link:
[https://coronavirus.politologue.com/data/coronavirus/coronacsv.aspx?format=json](https://coronavirus.politologue.com/data/coronavirus/coronacsv.aspx?format=json).