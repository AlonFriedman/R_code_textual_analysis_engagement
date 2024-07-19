# Load necessary libraries
library(readxl)
library(dplyr)
library(tidyr)
library(syuzhet)
library(stringr)
library(textstem)   # For stemming
library(tidytext)   # For tokenization and sentiment analysis
library(tm)         # For removeWords function

# Open file dialog to choose the Excel file
file_path <- file.choose()

# Load the dataset
Undergraduate2022_2019 <- read_excel(file_path)

# Print the dimensions of the dataset
cat("Dimensions of the dataset:\n")
print(dim(Undergraduate2022_2019))

# Print the first few rows
cat("\nFirst few rows of the dataset:\n")
print(head(Undergraduate2022_2019))

# Function to clean text data
clean_text <- function(text) {
  text <- tolower(text)
  text <- str_replace_all(text, "[[:punct:]]", " ")
  text <- str_replace_all(text, "[[:digit:]]", " ")
  text <- removeWords(text, stopwords("en"))
  text <- str_squish(text)
  text <- stem_words(text, language = "en")
  return(text)
}

# Print column names to identify the correct ones
print(colnames(Undergraduate2022_2019))

# Manually check and update these variables based on the actual column names in your dataset
project_col <- "Project.Type"  # Replace with the actual column name for project type
comment_col <- "Review.Comments"  # Replace with the actual column name for comments

# Verify and set the project column name
if (!(project_col %in% colnames(Undergraduate2022_2019))) {
  stop(paste("Column", project_col, "not found in the dataset. Please verify column names."))
}

# Verify and set the comment column name
if (!(comment_col %in% colnames(Undergraduate2022_2019))) {
  stop(paste("Column", comment_col, "not found in the dataset. Please verify column names."))
}

# Filter the data for final projects (adjust this if needed)
final_project <- Undergraduate2022_2019 %>%
  filter(grepl("Final", !!sym(project_col), ignore.case = TRUE))

# Check if final_project is correctly created
if (nrow(final_project) == 0) {
  stop("No records found for 'Final' projects. Please verify your data and column names.")
}

# Calculate the sample size
sample_size <- nrow(final_project)

# Print the sample size
cat("\nSample size:", sample_size, "\n")

# Function to analyze reviews with POS tagging and sentiment analysis
analyze_reviews <- function(data, comment_col) {
  # Clean the 'Comments' column
  data[[comment_col]] <- sapply(data[[comment_col]], clean_text)
  
  # Tokenization using tidytext
  data_tokens <- data %>%
    unnest_tokens(word, !!sym(comment_col))
  
  # Sentiment analysis using syuzhet
  sentiments <- get_nrc_sentiment(as.character(data[[comment_col]]))
  
  data <- data %>%
    mutate(
      words_per_review = str_count(data[[comment_col]], "\\w+"),
      sentiment_score = rowSums(sentiments[, c("positive", "negative")]),
      negative_score = sentiments$negative,
      positive_score = sentiments$positive
    )
  
  data_summary <- data %>%
    summarise(
      avg_words_per_review = mean(words_per_review, na.rm = TRUE),
      sd_words_per_review = sd(words_per_review, na.rm = TRUE),
      avg_sentiment_score = mean(sentiment_score, na.rm = TRUE),
      sd_sentiment_score = sd(sentiment_score, na.rm = TRUE),
      avg_negative_score = mean(negative_score, na.rm = TRUE),
      sd_negative_score = sd(negative_score, na.rm = TRUE),
      avg_positive_score = mean(positive_score, na.rm = TRUE),
      sd_positive_score = sd(positive_score, na.rm = TRUE),
      avg_negative_score_ratio = mean(negative_score / (negative_score + positive_score), na.rm = TRUE),
      sd_negative_score_ratio = sd(negative_score / (negative_score + positive_score), na.rm = TRUE)
    )
  
  return(data_summary)
}

# Analyze final project reviews
final_summary <- analyze_reviews(final_project, comment_col)

# Print final project summary and confidence intervals
cat("Final Project Summary:\n")
print(as.data.frame(final_summary))

# Function to calculate confidence intervals for negative keyword ratio
calc_conf_interval <- function(mean, sd, n) {
  error <- qnorm(0.975) * sd / sqrt(n)
  lower <- mean - error
  upper <- mean + error
  return(c(lower = lower, upper = upper))
}

# Calculate confidence intervals for final project
final_conf_interval <- calc_conf_interval(
  mean = final_summary$avg_negative_score_ratio,
  sd = final_summary$sd_negative_score_ratio,
  n = sample_size
)

cat("\n95% Confidence Interval for Negative Keyword Ratio:\n")
print(final_conf_interval)

# Print the first few reviews
cat("\nFirst few reviews:\n")
print(head(final_project[[comment_col]]))

# Print the number of reviews
cat("\nNumber of reviews:", sample_size, "\n")

# Print unique words in reviews
unique_words <- final_project[[comment_col]] %>%
  unlist() %>%
  strsplit(" ") %>%
  unlist() %>%
  unique()

cat("\nSample of unique words in reviews:\n")
print(head(unique_words, 20))
