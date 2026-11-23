# Bike Sharing Analytics — One Dataset, Four Technologies

**Status:** In Progress

**Tech Stack:** Python (pandas, Matplotlib) | R (dplyr, ggplot2) | Excel VBA | Power BI (Power Query, DAX)

## Overview

This project explores how the same data analytics problem can be solved using four different technologies.

Using the UCI Bike Sharing Dataset, the goal is to independently implement equivalent analytical workflows in Python, R, Excel VBA, and Power BI, and compare their outputs to evaluate numerical consistency and reproducibility.

The analysis focuses on identifying bike rental demand patterns across hours, working-day categories, weather conditions, and months.

## Project Objectives

- Perform data inspection, validation, and transformation using different technologies.
- Calculate consistent KPIs and grouped statistics from the same source dataset.
- Compare data manipulation and visualization approaches across Python, R, Excel, and Power BI.
- Automate analytical reporting using Excel VBA.
- Develop an interactive Power BI dashboard with DAX measures and filtering.
- Validate the numerical consistency of the analytical results across all implementations.

## Methodology

Each technology independently processes the original dataset and produces the same five standardized output tables:

| Output | Description |
|---|---|
| summary.csv | Total rentals and average hourly demand |
| hourly.csv | Average rentals by hour |
| workingday.csv | Average rentals by working-day category |
| weather.csv | Average rentals by weather category |
| monthly.csv | Monthly rental totals |

The results will be compared programmatically to identify discrepancies and verify analytical consistency.

## Dataset

**Source:** UCI Machine Learning Repository — Bike Sharing Dataset

**Reference**: [1] Fanaee-T, H. (2013). Bike Sharing [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5W894.

**Link:** https://archive.ics.uci.edu/dataset/275/bike+sharing+dataset

The dataset contains 17,379 recorded hourly observations from 2011–2012, including timestamps, rental counts, weather information, and working-day indicators.

## Current Status

The project is under active development. Individual implementations are being prepared and tested. Cross-technology validation, final dashboards, and analytical findings will be documented as they are completed.

## Expected Deliverables

- Reproducible Python notebook and R script
- Excel workbook with VBA reporting automation
- Interactive Power BI dashboard
- Standardized CSV outputs for each implementation
- Cross-technology comparison and validation
- Documented analytical findings and visualizations

## Key Principle

**One dataset. One analytical objective. Four independent implementations. Consistent, verifiable results.**