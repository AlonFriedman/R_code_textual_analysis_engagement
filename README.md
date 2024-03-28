This repository contains the R code for analyzing student peer reviews in an R class. 
This work, titled "A Case Study Using Textual Analysis to Infer Student Engagement in Online Data Science Education," will be published in the Journal of Statistics and Data Science Education. The analysis aims to uncover insights into student engagement through the lens of textual feedback provided during peer reviews.

Installation:
To run the code provided in this repository, you will need R installed on your machine along with the following R packages:
--haven
--dplyr
--tidytext
--textclean
--syuzhet
--ggplot2
--readxl 
--tm 
--and udpipe 

You can install these packages using the following command in your R console:
install.packages(c("haven", "dplyr", "tidytext", "textclean", "syuzhet", "ggplot2", "udpipe",  "readxl", "tm "))

Usage:
Clone or Download: Clone or download this repository to your local machine.
Install Packages: Install the required R packages listed in the Installation section if you haven't already (refer to the above command).
Prepare Dataset: Ensure you have an Excel file formatted according to the dataset structure described below. Update the filepath variable in the scripts with the path to your dataset file
filepath <- "path/to/your/Undergraduate2022-2019.xlsx" # Replace with the actual path to your dataset

Run Scripts: 
Open each R script in RStudio or your preferred R environment and execute them in the following order:

--R_Code_textual_analysis_engagement.R


Data:
The dataset required for running this code consists of 11 columns and 439 rows, including student final grades in the class, assignment grades, student comments on peer work, and a breakdown of rubric scores. No sample data is provided.
The structure of the dataset is as follows:

--Grade Received on Assignment
--Overall Grade in the Class
--Review Comments Given
--Overall
--The Statistics Contribution
--Code
--Documentation
--Testing
--Style
--Organization
--Accessibility

Contributing:
We welcome contributions to this project! If you have suggestions for improvements or have found bugs, please open an issue or submit a pull request. For major changes, please open an issue first to discuss what you would like to change.

License:
This work is distributed under the GPL license. This allows you to freely copy, modify, and distribute the code for personal use. However, if you publish modified versions or bundle it with other code, the modified version or complete bundle must also be licensed under the GPL.
