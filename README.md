# ComprehensiveExploration
This Python algorithm scans a given folder for files (PDF or DOCX) and extracts text from them. It then uses a custom-trained spaCy model to identify entities such as locations, types of work, or project details. It compares these identified entities against predefined lists of relevant terms stored in JSON files to categorize the contents of the folder based on certain criteria (like location, project type, etc.). If certain entities are missing or ambiguous, it prompts the user for clarification or additional information. Finally, it returns the categorized results as a dictionary.

You need to provide the location of a folder containing a list of folders, which group files where you want to extract information from (it works only for .pdf and .docx).
