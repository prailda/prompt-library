Playbook: R Data Science Tutorial

## Overview
Create a data science tutorial using an R markdown notebook.

## What’s Needed From User
- Link to a dataset (csv file attachment or kaggle link)
- Specific task to create a data science tutorial for

## Procedure
1. Download the dataset provided by the user.
-  If needed, download the dataset using the Kaggle CLI - you don't need any credentials for this
2. Create an R markdown notebook titled `data_science_tutorial.Rmd`.
3. Create a `tmp.Rmd` file for writing and saving intermediate code.
4. Create 5 main sections inside the `data_science_tutorial.Rmd` file and add code from the `tmp.Rmd` file containing the following:
- Dataset Statistics. Generate a statistical summary of the dataset.
- EDA (Exploratory Data Analysis). Create a bar chart and a scatter plot for the provided data.
- Train-test split. Split the data in an 80:20 ratio. Save the training and testing data.
- Training the machine learning model. Save the model once trained.
- Inference with the saved model. Load the saved model and evaluate its performance on the test set using the metric specified by the user.
5. Once the code is written, add a short explanation for each section.
6. Convert the R markdown notebook to HTML format
7. Send the final R markdown notebook, HTML file, saved model and testing data to the user.

## Specifications
1. Send the R markdown notebook and HTML file to the user.
2. Send the saved model and testing data to the user.

## Advice and Pointers
1. Do not re-install packages if already installed.
2. Sign in to RStudio is not required to complete this task.
3. Run the entire notebook after you add code for each section.

## Forbidden Actions
1. Do not overwrite the `data_science_tutorial.Rmd` file.