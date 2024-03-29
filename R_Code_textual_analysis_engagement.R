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
  
  counts <- analyze_comments(df$Comments)
  
  plot <- ggplot(counts, aes(x = upos, y = count, fill = upos)) +
    geom_bar(stat = "identity") +
    facet_wrap(~Year + ProjectType) +
    theme_minimal() +
    labs(title = "arts-of-speech usage per project showing the number of adjectives, adverbs, and
nouns per review for the first and final projects 2019-2021")
  
  print(plot) # Explicitly print the plot
}

# Example usage with a generic path
process_and_visualize_data('path/to/your/Undergraduate2022-2019.xlsx')

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
    labs(title = "Figure 5: Key word usage by overall grade. Stacked bars represent the numbers of positive,
negative, and negating words per review, averaged over three semesters", x = "Grade", y = "Sentiment Score")
}

# Generate the plot using the defined data frame (replace with your actual data processing)
sentiment_plot <- generate_sentiment_plot(student_data)
print(sentiment_plot) # Explicitly print the plot


# figure # 6

# Define a function to read and process data from an Excel file
generate_sentiment_by_grade_plot <- function(filepath) {
  df <- read_and_process_data(filepath)
  
  plot <- ggplot(df, aes(x = grade)) +
    geom_bar(aes(y = total_sentiment, fill = "Total Sentiment"), stat = "identity", position = "dodge") +
    geom_bar(aes(y = absolute_sentiment, fill = "Absolute Sentiment"), stat = "identity", position = "dodge", alpha = 0.5) +
    labs(title = "Sentiment by Overall Grade, illustrating the total (positive minus negative)
sentiment and absolute value of sentiment per overall course letter grade, accompanied by
standard deviation. This analysis is based on a dataset of 406 student reviews collected
over all three semesters", x = "Grade", y = "Sentiment Score") +
    scale_fill_manual(values = c("Total Sentiment" = "blue", "Absolute Sentiment" = "red")) +
    theme_minimal()
  
  print(plot) # Explicitly print the plot
}

# Example usage with a generic path
generate_sentiment_by_grade_plot('path/to/your/Undergraduate2022-2019.xlsx')
  
