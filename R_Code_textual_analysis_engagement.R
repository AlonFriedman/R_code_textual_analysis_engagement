# Load necessary libraries (use haven for Excel import)
library(haven)
library(dplyr)
library(tidytext)  # Check if loaded, reload if necessary (optional)
library(textclean)  # For text preprocessing
library(syuzhet)  # For sentiment analysis
library(ggplot2)  # For plotting (optional)
library(tm)    # Optional, not used in this revised code
library(readxl)  
library(udpipe)

# Install haven if not already installed
if (!require(haven)) install.packages("haven")

# Specify the file path (update as needed)
filepath <- "path/to/your/Undergraduate2022-2019.xlsx"

# ... rest of your script ...

# Figure # 4

# Download and load the udpipe model for English
ud_model <- udpipe_download_model(language = "english", model_dir = ".")
ud_model <- udpipe_load_model(ud_model$file_model)

# Function to analyze comments and count POS
analyze_comments <- function(comments) {
  annotated <- udpipe_annotate(ud_model, x = comments)
  df_annotated <- as.data.frame(annotated)
  
  counts <- df_annotated %>%
    group_by(upos) %>%
    summarise(count = n()) %>%
    filter(upos %in% c("ADJ", "ADV", "NOUN"))
  
  return(counts)
}

# Function to process data and visualize
process_and_visualize_data <- function(filepath) {
  df <- read_excel(filepath)
  
  # Assuming the DataFrame df has columns 'Year', 'ProjectType', and 'Comments'
  # Adjust column names based on your actual data
  counts <- analyze_comments(df$Comments)
  
  # Visualization
  ggplot(counts, aes(x = upos, y = count, fill = upos)) +
    geom_bar(stat = "identity") +
    facet_wrap(~Year + ProjectType) +
    theme_minimal() +
    labs(title = "POS Counts per Project Type and Year", x = "Part of Speech", y = "Count")
}

# Example usage
# Replace 'path_to_your_data.xlsx' with the actual path to your Excel file
process_and_visualize_data('path_to_your_data.xlsx')

#Figure 5.

# Define the data (replace with your actual data or function to process it)
student_data <- data.frame(
  grade = c("A", "B", "C", "D", "F"),  # Replace with actual grades
  sentiment_score = c(0.8, 0.5, 0.2, -0.1, -0.3),  # Replace with actual scores
  sentiment_label = c("positive", "positive", "neutral", "negative", "negative")  # Replace with labels
)

generate_sentiment_plot <- function(data) {
  # Assuming 'data' is a data frame with 'grade', 'sentiment_score', and 'sentiment_label' columns
  
  ggplot(data, aes(x = grade, y = sentiment_score, fill = sentiment_label)) +
    geom_bar(stat = "identity", position = "dodge") +
    theme_minimal() +
    labs(title = "Sentiment by Grade", x = "Grade", y = "Sentiment Score")
}

# Generate the plot using the defined data frame (replace with your actual data processing)
sentiment_plot <- generate_sentiment_plot(student_data)

# Save the plot to a file (optional)
ggsave("sentiment_plot.png", sentiment_plot, width = 8, height = 6)  # Adjust width and height as needed

# figure # 6

# Define a function to read and process data from an Excel file
read_and_process_data <- function(filepath) {
  # Read the Excel file
  student_data <- read_excel(filepath)
  
  # Assuming 'analyze_student_data' is a hypothetical function that would analyze the comments
  # and return a dataframe with 'grade',
  