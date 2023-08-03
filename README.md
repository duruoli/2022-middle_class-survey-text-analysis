# Middle-class Imagination Survey
**Keywords**: Survey, Text Analysis, Middle Class

## Abstract
Utilize 60 field survey questionnaires for text analysis. Apply clustering methods to identify common perceptions of the middle class among residents in Hangzhou.

## Data
Hangzhou "Thousand Talents Program" Field Survey

Question related to middle class:
1. How do you understand the concept of the middle class?
2. Who do you consider to be part of the middle class?
3. Do you consider yourself to be part of the middle class?
4. How do you perceive the future development of the middle class?
5. What role do you believe the middle class plays in society?
6. In your opinion, what are some unique characteristics of the middle class compared to other groups?
7. Have you heard of the term "middle class anxiety"?
8. Do you personally experience this kind of anxiety?
9. How would you describe the relationship between the middle class and the upper class?
10. How about the relationship with the lower class?

## Process
1. Data cleaning: transformation into structured dataset("normalization"), missing data imputation
2. Analysis: prediction (e.g. whether or not know the term of "middle class", accuracy: 86.67%); clustering

## Problems
1. Data size: Too small sample size (60), hard to build robust prediction models
2. Data missing: As this is a field survey, it is not guaranteed that all questions will be covered in each survey, leading to potential missing data; furthermore, being text data, proper imputation can be challenging.
