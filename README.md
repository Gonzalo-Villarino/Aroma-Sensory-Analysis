# Aroma Sensory Analysis Pipeline

This repository contains an R pipeline used to analyze aroma sensory panel data and generate breeder-friendly summaries and plots.

## Inputs

The analysis requires two Excel files:

* **Flower_Aroma_Sensory_Workshop.xlsx**
  Raw panelist sensory votes (CATA format).

* **Mapping_Flower_LM_Table_Codes_Secret.xlsx**
  Mapping table linking LM identifiers, codes, and GIDs.

## Outputs

The pipeline generates:

* `Max_Aroma_Summary.xlsx`
* `Panelist_Raw_Votes.xlsx`
* `Aroma_Category_Percentages.xlsx`
* LM_XXX_PG_CATA_plot.png
* LM_Percent_Profile.png plots in a subfolder

All outputs are exported automatically to the defined output directory.

## How to run

Open the script:

```
Aroma_Sensory_Analysis.R
```

Then run the pipeline in R or RStudio.

Required R packages include:

* tidyverse
* readxl
* writexl
* ggplot2
* dplyr

## Purpose

The goal of this pipeline is to summarize panel agreement on aroma categories and provide clear visual summaries for breeding and product development.
# Aroma-Sensory-Analysis
R pipeline for aroma sensory panel analysis and visualization
