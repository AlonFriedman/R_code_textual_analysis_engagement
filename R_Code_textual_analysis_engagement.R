# Calculate the standard error (SE) and margin of error (ME)
calculate_ci <- function(proportion, n) {
  se <- sqrt((proportion * (1 - proportion)) / n)
  me <- se * 1.96
  list(se = se, me = me)
}

# Apply CI calculation
sampled_data <- sampled_data %>%
  mutate(pos_ratio = pos_count / (pos_count + neg_count),
         ci = list(calculate_ci(pos_ratio, n()))) %>%
  unnest_wider(ci)

# Plot results
ggplot(sampled_data, aes(x = Grade, y = total_sentiment, fill = Grade)) +
  geom_bar(stat = "identity", position = "dodge") +
  geom_errorbar(aes(ymin = total_sentiment - me, ymax = total_sentiment + me), width = 0.2) +
  theme_minimal() +
  labs(title = "Total Sentiment by Grade",
       x = "Grade", y = "Total Sentiment") +
  scale_fill_manual(values = c("A" = "blue", "B" = "green", "C" = "yellow", "D" = "orange", "F" = "red"))

# Print sample summary
sampled_data %>%
  summarise(mean_pos = mean(pos_count), mean_neg = mean(neg_count),
            mean_negating = mean(negating_count), mean_sentiment = mean(total_sentiment))

# Additional analysis
summary_table <- sampled_data %>%
  group_by(Grade) %>%
  summarise(Reviews = n(), Pos = mean(pos_count), Neg = mean(neg_count),
            Total = mean(total_sentiment), Negate = mean(negating_count),
            Words_Per_Sentence = mean(str_count(Comments, "\\w+")),
            Words_Per_Review = mean(str_count(Comments, "\\s+") + 1),
            Adj = mean(ADJ), Adv = mean(ADV), Noun = mean(NOUN))

print(summary_table)

# Save the plot
ggsave("sentiment_by_grade.png")

# Save the summary table to a CSV file
write.csv(summary_table, "summary_table.csv", row.names = FALSE)
