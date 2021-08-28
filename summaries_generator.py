import docx
import os

# take in name of subject, & topics as a list
teaching_period = "21T4"
subject = "Applied Data Science"
topics = [
    "Data sources",
    "Noisy data & reliability",
    "Analysis & verification",
    "Visualisation",
    "Validation",
    "Presenting results for decision-making",
    "Revision",
]


# create a new folder with the name of subject
current_dir = os.getcwd()

final_dir = os.path.join(current_dir, rf"{teaching_period} {subject}")

if os.path.exists(final_dir):
    print("Error: folder already exists")
else:
    os.makedirs(final_dir)


# loop and title each topic as title formatting
for i in range(len(topics)):
    doc = docx.Document()

    doc_title = f"{i + 1} {topics[i]}"
    doc.add_paragraph(doc_title, "Title")

    doc.save(f"{final_dir}\\" + f"{doc_title}.docx")
