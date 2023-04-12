# Exam_Gen to Shuffle Answer Options

This repository contains a project to generate an exam with shuffled answer options from DOCX formatted file. 

## Purpose

When you create exams for students, making multiple form by shuffling answers are useful for teacher to prevent cheating. This project read questions, answers, and, answer keys from DOCX formatted exam file and create a new file with the same questions and shuffled answers.

<p align="center">
<img src="media/Shuffle Answers.png" width = 80% class="center">
</p>

## Usage

It requires python 3.9+ and libraries.

```python
from zipfile import ZipFile
from bs4 import BeautifulSoup
import unicodedata
import re
import pandas as pd
import random
```

Copy **Exam_Gen.py**  to the folder where contains original exam in docx format, run following commend in terminal

```
python Exam_Gen.py
```

It will prompt for filename and then create new DOCX file with the shuffled answers. The answer keys will be saved in separate CSV file.
