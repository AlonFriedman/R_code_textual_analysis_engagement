# Load necessary libraries
library(readxl)
library(dplyr)
library(tidyr)
library(stringr)
library(tidytext)
library(ggplot2)
library(knitr)
library(tibble)

# Function to safely read the dataset
read_data <- function(file_path) {
  tryCatch({
    data <- read_excel(file_path)
    cat("Data loaded successfully.\n")
    return(data)
  }, error = function(e) {
    cat("Error in reading the Excel file: ", e$message, "\n")
    stop("Failed to load data.")
  })
}

# Load the data
cat("Please select your Excel file in the file dialog that opens.\n")
file_path <- file.choose()
projects <- read_data(file_path)

# Add an ID column
projects <- projects %>% rownames_to_column(var = "ID")

# Print the dimensions and first few rows of the dataset
cat("Dimensions of the dataset:\n")
print(dim(projects))
cat("\nFirst few rows of the dataset:\n")
print(head(projects))

# Function to convert letter grades to numeric values
convert_grade_to_numeric <- function(grade) {
  grade_scale <- c(
    "A+" = 100, "A" = 95, "A-" = 90,
    "B+" = 88, "B" = 85, "B-" = 80,
    "C+" = 78, "C" = 75, "C-" = 70,
    "D+" = 68, "D" = 65, "D-" = 60,
    "F" = 40
  )
  grade <- as.character(grade)
  grade <- trimws(grade)
  numeric_grade <- grade_scale[grade]
  ifelse(is.na(numeric_grade), NA_real_, numeric_grade)
}

# Apply conversion
projects <- projects %>%
  mutate(
    First_Grade_Numeric = convert_grade_to_numeric(`First assignment grade`),
    Final_Grade_Numeric = convert_grade_to_numeric(`Final assignment grade`)
  )

# Rename columns for clarity
projects <- projects %>%
  rename(
    First_Grade = `First assignment grade`,
    Final_Grade = `Final assignment grade`,
    Overall_Grade = `Overall grade`,
    Review_Comments = `Review Comments Given`,
    Comments = `Overal grade for the final project`
  )

# Function to clean text data
clean_text <- function(text) {
  text <- tolower(text)
  text <- str_replace_all(text, "[[:punct:]]", " ")
  text <- str_replace_all(text, "[[:digit:]]", " ")
  text <- str_squish(text)
  return(text)
}

# Process the text data
projects <- projects %>%
  mutate(
    cleaned_comment = clean_text(as.character(Comments)),
    word_count = str_count(as.character(Comments), "\\S+"),
    sentence_count = str_count(as.character(Comments), "[.!?]+"),
    words_per_sentence = ifelse(sentence_count > 0, word_count / sentence_count, NA),
    Grade_Improvement = Final_Grade_Numeric - First_Grade_Numeric
  )

# Load sentiment lexicon
sentiment_lexicon <- get_sentiments("bing")

# Add Negate words
negate_words <- c("not", "no", "never", "without", "any")
negative_words <- c("bad", "poor", "second-rate", "Something work", "do not understand", 
                    "not sure", "none", "confusing", "problem", "problems", "error", 
                    "errors", "did not see", "redo", "redoing", "lack", "unacceptable",
                    "mistake", "mistakes", "issue", "issues", "more", "waste", "unfinished",
                    "why", "unable", "taken", "dislike", "failed")
positive_words <- c("great", "outstanding", "excellent", "perfect", "good", "simplistic", 
                    "efficient", "interesting", "unique", "like", "understand", "nice",
                    "strong", "well", "done", "enjoy", "learn", "learned", "simple",
                    "useful", "creative", "works", "sense", "clear", "sharp", "easy", 
                    "complete", "documentation")

# Define basic part of speech lists
adjectives <- c("good", "great", "bad", "nice", "excellent", "poor", "wonderful", "terrible", "amazing", "awful")
adverbs <- c("very", "really", "quite", "extremely", "fairly", "pretty", "too", "almost", "enough", "hardly")
nouns <- c("student", "assignment", "project", "work", "essay", "paper", "research", "analysis", "report", "study")

sentiment_lexicon <- rbind(sentiment_lexicon, data.frame(word = negate_words, sentiment = "Negate"))

# Tokenize comments and perform sentiment analysis
comment_sentiments <- projects %>%
  select(id = ID, text = cleaned_comment) %>%
  unnest_tokens(word, text) %>%
  mutate(
    word_index = row_number(),
    prev_word = lag(word),
    next_word = lead(word)
  ) %>%
  mutate(
    is_positive = word %in% positive_words,
    is_negative = word %in% negative_words,
    is_negated = (prev_word %in% negate_words) | (next_word %in% negate_words)
  ) %>%
  group_by(id) %>%
  summarise(
    positive = sum(is_positive & !is_negated, na.rm = TRUE),
    negative = sum(is_negative & !is_negated, na.rm = TRUE) + sum(is_positive & is_negated, na.rm = TRUE),
    negate = sum(word %in% negate_words, na.rm = TRUE),
    noun_count = sum(word %in% nouns, na.rm = TRUE),
    adjective_count = sum(word %in% adjectives, na.rm = TRUE),
    adverb_count = sum(word %in% adverbs, na.rm = TRUE),
    total_words = n(),
    sentiment_score = (positive - negative) / total_words,
    sentiment_category = case_when(
      sentiment_score > 0 ~ "Positive",
      sentiment_score < 0 ~ "Negative",
      TRUE ~ "Negate"
    )
  )

# Join sentiment analysis results back to the main dataset
projects <- projects %>%
  left_join(comment_sentiments, by = c("ID" = "id"))

# Create summary tables for the entire dataset
overall_summary <- projects %>%
  summarise(
    Total_Reviews = n(),
    Avg_Words_Per_Review = mean(word_count, na.rm = TRUE),
    Avg_Sentences_Per_Review = mean(sentence_count, na.rm = TRUE),
    Avg_Words_Per_Sentence = mean(words_per_sentence, na.rm = TRUE),
    Avg_First_Grade = mean(First_Grade_Numeric, na.rm = TRUE),
    Avg_Final_Grade = mean(Final_Grade_Numeric, na.rm = TRUE),
    Avg_Grade_Improvement = mean(Grade_Improvement, na.rm = TRUE),
    Positive_Reviews = sum(sentiment_category == "Positive", na.rm = TRUE),
    Negative_Reviews = sum(sentiment_category == "Negative", na.rm = TRUE),
    Negate_Reviews = sum(sentiment_category == "Negate", na.rm = TRUE),
    Avg_Positive_Words = mean(positive, na.rm = TRUE),
    Avg_Negative_Words = mean(negative, na.rm = TRUE),
    Avg_Negate_Words = mean(negate, na.rm = TRUE),
    Avg_Nouns = mean(noun_count, na.rm = TRUE),
    Avg_Adjectives = mean(adjective_count, na.rm = TRUE),
    Avg_Adverbs = mean(adverb_count, na.rm = TRUE)
  )

# Print overall summary
cat("\nOverall Summary:\n")
print(kable(overall_summary, format = "pipe", digits = 2))

# Summary by year
year_summary <- projects %>%
  group_by(Year) %>%
  summarise(
    Total_Reviews = n(),
    Avg_Words_Per_Review = mean(word_count, na.rm = TRUE),
    Avg_Sentences_Per_Review = mean(sentence_count, na.rm = TRUE),
    Avg_Words_Per_Sentence = mean(words_per_sentence, na.rm = TRUE),
    Avg_First_Grade = mean(First_Grade_Numeric, na.rm = TRUE),
    SD_First_Grade = sd(First_Grade_Numeric, na.rm = TRUE),
    Avg_Final_Grade = mean(Final_Grade_Numeric, na.rm = TRUE),
    SD_Final_Grade = sd(Final_Grade_Numeric, na.rm = TRUE),
    Avg_Grade_Improvement = mean(Grade_Improvement, na.rm = TRUE),
    Positive_Reviews = sum(sentiment_category == "Positive", na.rm = TRUE),
    Negative_Reviews = sum(sentiment_category == "Negative", na.rm = TRUE),
    Negate_Reviews = sum(sentiment_category == "Negate", na.rm = TRUE),
    Avg_Positive_Words = mean(positive, na.rm = TRUE),
    Avg_Negative_Words = mean(negative, na.rm = TRUE),
    Avg_Negate_Words = mean(negate, na.rm = TRUE),
    Avg_Nouns = mean(noun_count, na.rm = TRUE),
    Avg_Adjectives = mean(adjective_count, na.rm = TRUE),
    Avg_Adverbs = mean(adverb_count, na.rm = TRUE)
  )

# Print year summary
cat("\nSummary by Year:\n")
print(kable(year_summary, format = "pipe", digits = 2))

# Summary by overall grade
grade_summary <- projects %>%
  group_by(Overall_Grade) %>%
  summarise(
    Total_Reviews = n(),
    Avg_Words_Per_Review = mean(word_count, na.rm = TRUE),
    Avg_Sentences_Per_Review = mean(sentence_count, na.rm = TRUE),
    Avg_Words_Per_Sentence = mean(words_per_sentence, na.rm = TRUE),
    Avg_First_Grade = mean(First_Grade_Numeric, na.rm = TRUE),
    Avg_Final_Grade = mean(Final_Grade_Numeric, na.rm = TRUE),
    Avg_Grade_Improvement = mean(Grade_Improvement, na.rm = TRUE),
    Positive_Reviews = sum(sentiment_category == "Positive", na.rm = TRUE),
    Negative_Reviews = sum(sentiment_category == "Negative", na.rm = TRUE),
    Negate_Reviews = sum(sentiment_category == "Negate", na.rm = TRUE),
    Avg_Positive_Words = mean(positive, na.rm = TRUE),
    Avg_Negative_Words = mean(negative, na.rm = TRUE),
    Avg_Negate_Words = mean(negate, na.rm = TRUE),
    Avg_Nouns = mean(noun_count, na.rm = TRUE),
    Avg_Adjectives = mean(adjective_count, na.rm = TRUE),
    Avg_Adverbs = mean(adverb_count, na.rm = TRUE)
  ) %>%
  arrange(factor(Overall_Grade, levels = c("A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-", "D+", "D", "D-", "F")))

# Print grade summary
cat("\nSummary by Overall Grade:\n")
print(kable(grade_summary, format = "pipe", digits = 2))

# Define a single, robust calc_ci function
calc_ci <- function(x) {
  x <- x[is.finite(x)]  # Remove NaN and Inf values
  if (length(x) < 2) {
    return(c(lower = NA, upper = NA))
  }
  se <- sd(x, na.rm = TRUE) / sqrt(length(x))
  mean_x <- mean(x, na.rm = TRUE)
  ci <- qt(0.975, df = length(x) - 1) * se
  return(c(lower = mean_x - ci, upper = mean_x + ci))
}

# Calculate confidence intervals
ci_summary <- projects %>%
  summarise(
    words_ci = list(calc_ci(word_count)),
    sentences_ci = list(calc_ci(sentence_count)),
    words_per_sentence_ci = list(calc_ci(words_per_sentence[is.finite(words_per_sentence)])),
    first_grade_ci = list(calc_ci(First_Grade_Numeric)),
    final_grade_ci = list(calc_ci(Final_Grade_Numeric)),
    grade_improvement_ci = list(calc_ci(Grade_Improvement)),
    sentiment_score_ci = list(calc_ci(sentiment_score))
  ) %>%
  gather(key = "metric", value = "ci") %>%
  mutate(
    lower = sapply(ci, `[`, 1),
    upper = sapply(ci, `[`, 2)
  ) %>%
  select(-ci)

# Print confidence intervals
cat("\nConfidence Intervals:\n")
print(kable(ci_summary, format = "pipe", digits = 2))

# Sample table
cat("\nSample of processed data (first 10 rows):\n")
sample_data <- projects %>%
  select(Year, Overall_Grade, First_Grade, Final_Grade, First_Grade_Numeric, Final_Grade_Numeric, 
         Grade_Improvement, word_count, sentence_count, words_per_sentence, 
         sentiment_score, sentiment_category) %>%
  head(10)
print(kable(sample_data, format = "pipe", digits = 2))

# Confidence intervals based on grades
ci_by_grade <- projects %>%
  group_by(Overall_Grade) %>%
  summarise(
    First_Grade_CI = list(calc_ci(First_Grade_Numeric)),
    First_Grade_SE = (sd(First_Grade_Numeric)/sqrt(length(First_Grade_Numeric))),
    First_Grade_ME = (1.96 * First_Grade_SE),
    Final_Grade_CI = list(calc_ci(Final_Grade_Numeric)),
    Final_Grade_SE = (sd(Final_Grade_Numeric)/sqrt(length(Final_Grade_Numeric))),
    Final_Grade_ME = (1.96 * Final_Grade_SE),
    Grade_Improvement_CI = list(calc_ci(Grade_Improvement)),
    Grade_Improvement_SE = (sd(Grade_Improvement)/sqrt(length(Grade_Improvement))),
    Grade_Improvement_ME = (1.96 * Grade_Improvement_SE)
  ) %>%
  hoist(First_Grade_CI, First_Grade_CI_Lower = 1, First_Grade_CI_Upper = 2) %>%
  hoist(Final_Grade_CI, Final_Grade_CI_Lower = 1, Final_Grade_CI_Upper = 2) %>%
  hoist(Grade_Improvement_CI, Grade_Improvement_Lower = 1, Grade_Improvement_Upper = 2)

cat("\nConfidence Intervals by Overall Grade:\n")
print(kable(ci_by_grade, format = "pipe", digits = 2))

# Confidence intervals based on year
ci_by_year <- projects %>%
  group_by(Year) %>%
  summarise(
    First_Grade_CI = list(calc_ci(First_Grade_Numeric)),
    First_Grade_SE = (sd(First_Grade_Numeric)/sqrt(length(First_Grade_Numeric))),
    First_Grade_ME = (1.96 * First_Grade_SE),
    Final_Grade_CI = list(calc_ci(Final_Grade_Numeric)),
    Final_Grade_SE = (sd(Final_Grade_Numeric)/sqrt(length(Final_Grade_Numeric))),
    Final_Grade_ME = (1.96 * Final_Grade_SE),
    Grade_Improvement_CI = list(calc_ci(Grade_Improvement)),
    Grade_Improvement_SE = (sd(Grade_Improvement)/sqrt(length(Grade_Improvement))),
    Grade_Improvement_ME = (1.96 * Grade_Improvement_SE)
  ) %>%
  hoist(First_Grade_CI, First_Grade_CI_Lower = 1, First_Grade_CI_Upper = 2) %>%
  hoist(Final_Grade_CI, Final_Grade_CI_Lower = 1, Final_Grade_CI_Upper = 2) %>%
  hoist(Grade_Improvement_CI, Grade_Improvement_Lower = 1, Grade_Improvement_Upper = 2)

cat("\nConfidence Intervals by Year:\n")
print(kable(ci_by_year, format = "pipe", digits = 2))

# Sample a portion of the data
total_size <- nrow(projects)
sample_size <- 100

set.seed(123)
sample_100 <- projects %>%
  slice_sample(n = sample_size)

# Function to calculate margin of error
calc_margin_of_error <- function(x, confidence = 0.95) {
  x <- x[is.finite(x)]
  n <- length(x)
  if (n < 2) return(NA)
  
  se <- sd(x) / sqrt(n)
  t_score <- qt((1 + confidence) / 2, df = n - 1)
  margin_of_error <- t_score * se
  
  return(margin_of_error)
}

margin_of_error <- qnorm(0.975) * sqrt(0.5 * 0.5 / sample_size) * sqrt((total_size - sample_size) / (total_size - 1))
cat("Margin of Error:", round(margin_of_error * 100, 2), "%\n")

# Calculate summary statistics and margin of error
summary_stats <- sample_100 %>%
  summarise(
    Sample_Size = n(),
    Avg_First_Grade = mean(First_Grade_Numeric, na.rm = TRUE),
    Avg_Final_Grade = mean(Final_Grade_Numeric, na.rm = TRUE),
    Avg_Grade_Improvement = mean(Grade_Improvement, na.rm = TRUE),
    ME_First_Grade = calc_margin_of_error(First_Grade_Numeric),
    ME_Final_Grade = calc_margin_of_error(Final_Grade_Numeric),
    ME_Grade_Improvement = calc_margin_of_error(Grade_Improvement)
  )

# Print summary statistics and margin of error
cat("\nSummary Statistics and Margin of Error (95% Confidence Level):\n")
print(kable(summary_stats, format = "pipe", digits = 2))

# Calculate confidence intervals
sample_ci_summary <- sample_100 %>%
  summarise(
    words_ci = list(calc_ci(word_count)),
    sentences_ci = list(calc_ci(sentence_count)),
    words_per_sentence_ci = list(calc_ci(words_per_sentence[is.finite(words_per_sentence)])),
    first_grade_ci = list(calc_ci(First_Grade_Numeric)),
    final_grade_ci = list(calc_ci(Final_Grade_Numeric)),
    grade_improvement_ci = list(calc_ci(Grade_Improvement)),
    sentiment_score_ci = list(calc_ci(sentiment_score))
  ) %>%
  gather(key = "metric", value = "ci") %>%
  mutate(
    lower = sapply(ci, `[`, 1),
    upper = sapply(ci, `[`, 2)
  ) %>%
  select(-ci)

# Print confidence intervals
cat("\nSample Confidence Intervals:\n")
print(kable(sample_ci_summary, format = "pipe", digits = 2))

# Confidence intervals based on grades
sample_ci_by_grade <- sample_100 %>%
  group_by(Overall_Grade) %>%
  summarise(
    First_Grade_CI = list(calc_ci(First_Grade_Numeric)),
    First_Grade_SE = (sd(First_Grade_Numeric)/sqrt(length(First_Grade_Numeric))),
    First_Grade_ME = (1.96 * First_Grade_SE),
    Final_Grade_CI = list(calc_ci(Final_Grade_Numeric)),
    Final_Grade_SE = (sd(Final_Grade_Numeric)/sqrt(length(Final_Grade_Numeric))),
    Final_Grade_ME = (1.96 * Final_Grade_SE),
    Grade_Improvement_CI = list(calc_ci(Grade_Improvement)),
    Grade_Improvement_SE = (sd(Grade_Improvement)/sqrt(length(Grade_Improvement))),
    Grade_Improvement_ME = (1.96 * Grade_Improvement_SE)
  ) %>%
  hoist(First_Grade_CI, First_Grade_CI_Lower = 1, First_Grade_CI_Upper = 2) %>%
  hoist(Final_Grade_CI, Final_Grade_CI_Lower = 1, Final_Grade_CI_Upper = 2) %>%
  hoist(Grade_Improvement_CI, Grade_Improvement_Lower = 1, Grade_Improvement_Upper = 2)

cat("\nSample Confidence Intervals by Overall Grade:\n")
print(kable(sample_ci_by_grade, format = "pipe", digits = 2))

# Confidence intervals based on year
sample_ci_by_year <- sample_100 %>%
  group_by(Year) %>%
  summarise(
    First_Grade_CI = list(calc_ci(First_Grade_Numeric)),
    First_Grade_SE = (sd(First_Grade_Numeric)/sqrt(length(First_Grade_Numeric))),
    First_Grade_ME = (1.96 * First_Grade_SE),
    Final_Grade_CI = list(calc_ci(Final_Grade_Numeric)),
    Final_Grade_SE = (sd(Final_Grade_Numeric)/sqrt(length(Final_Grade_Numeric))),
    Final_Grade_ME = (1.96 * Final_Grade_SE),
    Grade_Improvement_CI = list(calc_ci(Grade_Improvement)),
    Grade_Improvement_SE = (sd(Grade_Improvement)/sqrt(length(Grade_Improvement))),
    Grade_Improvement_ME = (1.96 * Grade_Improvement_SE)
  ) %>%
  hoist(First_Grade_CI, First_Grade_CI_Lower = 1, First_Grade_CI_Upper = 2) %>%
  hoist(Final_Grade_CI, Final_Grade_CI_Lower = 1, Final_Grade_CI_Upper = 2) %>%
  hoist(Grade_Improvement_CI, Grade_Improvement_Lower = 1, Grade_Improvement_Upper = 2)

cat("\nSample Confidence Intervals by Year:\n")
print(kable(sample_ci_by_year, format = "pipe", digits = 2))

# Time Analysis
time_analysis <- projects %>%
  group_by(Year) %>%
  summarise(
    Number_Reviewers = n(),
    Positive_Reviews = sum(sentiment_category == "Positive"),
    Negative_Reviews = sum(sentiment_category == "Negative"),
    Negate_Reviews = sum(sentiment_category == "Negate"),
    Words_Per_Sentence = sum(words_per_sentence, na.rm = TRUE),
    Words_Per_Review = mean(word_count),
    Noun_Count = sum(noun_count),
    Adjective_Count = sum(adjective_count),
    Adverb_Count = sum(adverb_count)
  )

cat("\nTime Analysis:\n")
print(kable(time_analysis, format = "pipe", digits = 2))

sample_time_analysis <- sample_100 %>%
  group_by(Year) %>%
  summarise(
    Number_Reviewers = n(),
    Positive_Reviews = sum(sentiment_category == "Positive"),
    Negative_Reviews = sum(sentiment_category == "Negative"),
    Negate_Reviews = sum(sentiment_category == "Negate"),
    Words_Per_Sentence = sum(words_per_sentence, na.rm = TRUE),
    Words_Per_Review = mean(word_count),
    Noun_Count = sum(noun_count),
    Adjective_Count = sum(adjective_count),
    Adverb_Count = sum(adverb_count)
  )

cat("\nSample Time Analysis:\n")
print(kable(sample_time_analysis, format = "pipe", digits = 2))

metrics <- c("Positive_Reviews", "Negative_Reviews", "Words_Per_Sentence", 
             "Words_Per_Review", "Noun_Count", "Adjective_Count", "Adverb_Count")

anova_results <- list()

for (metric in metrics) {
  formula <- as.formula(paste(metric, "~ factor(Year)"))
  anova_result <- aov(formula, data = sample_time_analysis)
  anova_results[[metric]] <- summary(anova_result)
}

# Print ANOVA results
cat("\nANOVA Results:\n")
for (metric in names(anova_results)) {
  cat("\nANOVA for", metric, ":\n")
  print(anova_results[[metric]])
}
# Visualizations
# Sentiment distribution
ggplot(projects, aes(x = sentiment_category, fill = sentiment_category)) +
  geom_bar() +
  theme_minimal() +
  scale_fill_manual(values = c("Positive" = "green", "Negative" = "red", "Negate" = "gray")) +
  labs(title = "Distribution of Sentiment Categories", x = "Sentiment Category", y = "Count")
ggsave("sentiment_distribution.png", width = 8, height = 6)

# Visualizations for sentiment word counts
sentiment_data <- comment_sentiments %>%
  select(id, positive, negative, negate) %>%
  gather(key = "sentiment", value = "count", -id)

ggplot(sentiment_data, aes(x = sentiment, y = count)) +
  geom_boxplot() +
  theme_minimal() +
  labs(title = "Distribution of Sentiment Words", x = "Sentiment", y = "Count")
ggsave("sentiment_word_counts.png", width = 8, height = 6)

# Visualizations for basic part of speech counts
pos_data <- comment_sentiments %>%
  select(id, noun_count, adjective_count, adverb_count) %>%
  gather(key = "pos", value = "count", -id)

ggplot(pos_data, aes(x = pos, y = count)) +
  geom_boxplot() +
  theme_minimal() +
  labs(title = "Distribution of Basic Parts of Speech", x = "Part of Speech", y = "Count")
ggsave("basic_pos_distribution.png", width = 8, height = 6)

# Grade progression
ggplot(projects, aes(x = First_Grade_Numeric, y = Final_Grade_Numeric, color = sentiment_category)) +
  geom_point(alpha = 0.5) +
  geom_abline(intercept = 0, slope = 1, color = "red", linetype = "dashed") +
  theme_minimal() +
  labs(title = "Grade Progression: First Grade vs Final Grade", 
       x = "First Grade", 
       y = "Final Grade",
       subtitle = paste("Based on", sum(!is.na(projects$First_Grade_Numeric) & !is.na(projects$Final_Grade_Numeric)), 
                        "out of", nrow(projects), "total records")) +
  scale_color_manual(values = c("Negative" = "red", "Negate" = "gray", "Positive" = "green"))
ggsave("grade_progression.png", width = 8, height = 6)

# Grade improvement distribution
ggplot(projects, aes(x = Grade_Improvement)) +
  geom_histogram(binwidth = 1, fill = "lightblue", color = "black") +
  theme_minimal() +
  labs(title = "Distribution of Grade Improvement", x = "Grade Improvement", y = "Count")
ggsave("grade_improvement_distribution.png", width = 8, height = 6)
