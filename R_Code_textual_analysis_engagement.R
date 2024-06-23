# Load necessary libraries (use haven for Excel import)
library(haven)
library(dplyr)
library(tidytext)  # Check if loaded, reload if necessary (optional)
library(textclean)  # For text preprocessing
library(syuzhet)  # For sentiment analysis
library(ggplot2)  # For plotting (optional)
library(tm)    # Optional, not used in this revised code
library(readxl)
library(udpipe)

# Specify the file path (update as needed)
filepath <- "/path/to/your/"

# Load the pre-trained UDPipe model for English (update the path to your model)
ud_model <- udpipe_load_model("path/to/english-ud-2.0-170801.udpipe")

# Function to conduct sampling before analysis
conduct_sampling <- function(df, sample_size) {
  set.seed(123)  # For reproducibility
  sample_df <- df %>% sample_n(sample_size)
  return(sample_df)
}

# Function to calculate standard error and margin of error
calculate_statistics <- function(df, column_name) {
  p <- mean(df[[column_name]])
  n <- nrow(df)
  se <- sqrt(p * (1 - p) / n)
  me <- se * 1.96  # 95% confidence level
  return(list(SE = se, ME = me))
}

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
process_and_visualize_data <- function(filepath, sample_size) {
  df <- read_excel(filepath)
  sample_df <- conduct_sampling(df, sample_size)
  
  # Calculate statistics for sentiment analysis (assuming 'sentiment_score' column exists)
  stats <- calculate_statistics(sample_df, 'sentiment_score')
  print(paste("Standard Error (SE):", stats$SE))
  print(paste("Margin of Error (ME):", stats$ME))
  
  counts <- analyze_comments(sample_df$Comments)
  
  plot <- ggplot(counts, aes(x = upos, y = count, fill = upos)) +
    geom_bar(stat = "identity") +
    facet_wrap(~Year + ProjectType) +
    theme_minimal() +
    labs(title = "Parts-of-speech usage per project showing the number of adjectives, adverbs, and nouns per review for the first and final projects 2019-2021")
  
  print(plot) # Explicitly print the plot
}

# Example usage with a generic path and a sample size of 100
process_and_visualize_data('/path/to/your/', 100)

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
    labs(title = "Figure 5: Key word usage by overall grade. Stacked bars represent the numbers of positive, negative, and negating words per review, averaged over three semesters", x = "Grade", y = "Sentiment Score")
}

# Generate the plot using the defined data frame (replace with your actual data processing)
sentiment_plot <- generate_sentiment_plot(student_data)
print(sentiment_plot) # Explicitly print the plot

# Define a function to read and process data from an Excel file
generate_sentiment_by_grade_plot <- function(filepath) {
  df <- read_excel(filepath)
  
  plot <- ggplot(df, aes(x = grade)) +
    geom_bar(aes(y = total_sentiment, fill = "Total Sentiment"), stat = "identity", position = "dodge") +
    geom_bar(aes(y = absolute_sentiment, fill = "Absolute Sentiment"), stat = "identity", position = "dodge", alpha = 0.5) +
    labs(title = "Sentiment by Overall Grade, illustrating the total (positive minus negative) sentiment and absolute value of sentiment per overall course letter grade, accompanied by standard deviation. This analysis is based on a dataset of 406 student reviews collected over all three semesters", x = "Grade", y = "Sentiment Score") +
    scale_fill_manual(values = c("Total Sentiment" = "blue", "Absolute Sentiment" = "red")) +
    theme_minimal()
  
  print(plot) # Explicitly print the plot
}

# Example usage with a generic path
generate_sentiment_by_grade_plot('/path/to/your/')

