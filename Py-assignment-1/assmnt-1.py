from docx import Document
import re

"""
The purpose of this code is to go through a word document and count the instances of each word.
The words will then be sorted in ascending order of appearances.

Author: Stefan Saladino (200551988)
Inception Date: June 9, 2023
"""

# Main program
# sample.txt will be referred to as sample_filename
# report.txt will be referred to as report_filename
sample_filename = "sample.txt"
report_filename = "report.txt"

def count_words(sample_filename):
    # Open the file as read only
    with open(sample_filename, "r") as file:
        #sample.txt is being made case insensitive and being split at each word. Hyphenated words count as one.
        text = file.read().lower().split()

    # Count word occurrences using a dictionary
    word_count = {}
    #looping over each word in the sample doc
    for word in text:
        # Remove trailing punctuation from the word
        # substituting any character in a word that is not a word character, a hyphen, or an apostrophe for an empty string
        word = re.sub(r'[^\w\'-]', '', word)
        #if the word is new, set count to 1
        #if word has already appeared, add 1 to the current count
        if word in word_count:
            word_count[word] += 1
        else:
            word_count[word] = 1


    return word_count

    # Sort words by frequency in ascending order
def save_report(word_count, report_filename):
    sorted_words = sorted(word_count.items(), key=lambda x: x[1])
    
    longest_word_length = max(len(word) for word, _ in sorted_words)+1

    with open(report_filename, 'w') as report_file:
        header = '{:<{}} {:<2s}'.format('Word', longest_word_length, 'Frequency')
        underline = '-' * len(header)
        report_file.write(f"{header}\n{underline}\n\n")
        
        for word, count in sorted_words:
            report_file.write('{:<{}} ----> {:<2d}\n'.format(word+':', longest_word_length, count))


    # Count words and save the report
word_count = count_words(sample_filename)
save_report(word_count, report_filename)