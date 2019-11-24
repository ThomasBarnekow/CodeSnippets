#
# openxml_client.py
#
# Copyright 2019 Thomas Barnekow
#
# Developer: Thomas Barnekow
# Email: thomas<at/>barnekow<dot/>info

import clr
import shutil

clr.AddReference(
    r"..\CodeSnippets.OpenXmlWrapper\bin\Debug\net471\CodeSnippets.OpenXmlWrapper")

from CodeSnippets.OpenXmlWrapper import OpenXmlPowerToolsWrapper

wrapper = OpenXmlPowerToolsWrapper()

# Display contents before finishing review, showing that the document contains revision markup.
xml_with_revision_markup = wrapper.GetMainDocumentPart("DocumentWithRevisionMarkup.docx")
print("Document before finishing review:\n")
print(xml_with_revision_markup)

# Finish review, removing all revision markup.
print("\nFinishing review ...")
shutil.copyfile("DocumentWithRevisionMarkup.docx", "Result.docx")
xml_without_revision_markup = wrapper.FinishReview("Result.docx")

# Display contents after finishing review, showing that the revision markup was removed.
print("\nDocument after finishing review:\n")
print(xml_without_revision_markup)
