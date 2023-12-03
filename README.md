# Exam_gen: Create Shuffled Test Forms

## Overview

**Exam_gen** is a Python script designed to streamline the process of creating diverse sets of tests with shuffled answer orders. The script takes a Docx file containing questions and answers, shuffles the answer choices for each question, and generates multiple test forms with distinct answer orders.

### Key Features

- **Shuffle Answers:** Easily shuffle the order of answer choices for each question.
- **Form Generation:** Create multiple versions of tests with varied answer orders.
- **Key Information:** Retrieve key information for each generated test form.

### Use Cases

- Educational institutions promoting fair and unbiased assessments.
- Instructors seeking to discourage cheating by providing unique versions of exams.

## Dependencies
1. Downlolad and install Python 3.9+ from here: https://www.python.org/
2. To use **Exam_gen**, ensure you have the following dependencies installed:
    - [BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/) (version 4.12.2)
    - [pandas](https://pandas.pydata.org/) (version 2.1.1)
    - [numpy](https://numpy.org/) (version 1.26.0)

    Install the dependencies using:

```bash
pip install beautifulsoup4==4.12.2 pandas==2.1.1 numpy==1.26.0
```

## How to Use

Follow these steps to use **Exam_gen** and generate shuffled test forms:

1. Save the script `util.py` and `exam_gen.py` in the same directory as the DOCX file you want to convert.
2. Open a terminal window and navigate to the directory containing the scripts.
3. Execute the script using the following command:

    ```bash
    python exam_gen.py
    ```

4. Follow the on-screen prompts to provide the DOCX filename and specify the number of forms to generate.
5. The script will generate new DOCX files and a CSV file including answer keys in the new directory.

## Functionality

**Exam_gen** offers the following key functionalities:

### 1. Shuffling Answer Choices

The script shuffles the order of answer choices for each question in your DOCX file, providing a randomized distribution.

### 2. Form Generation

Multiple versions of tests (forms) are generated, each with a unique order of answer choices. This variation helps discourage cheating and ensures fair assessments.

### 3. Key Information Retrieval

The script extracts key information, providing a CSV file with answer keys for each generated test form.
Briefly describe the major functions in the code and their purposes.

## Input File Format

Ensure your DOCX file follows the expected format for successful processing by **Exam_gen**.

- The file should contain questions and answer choices formatted as paragraphs. Each question should be identified with a numerical index, and answer choices should be labeled with `A.`, `B.`, `C.`, and `D.`.

    *Example*
    ```
    1. What is the capital of France?
    A. London
    B. Berlin
    C. Paris
    D. Madrid
    ```

- **Exam_gen** identifies the form number by searching for the string "FORM X" within the DOCX file. It replaces the 'X' with the appropriate form letter from 'A' to 'H'. Ensure that your original document includes the form identifier in the format "FORM X" to enable accurate form numbering.
    
    *Example*
    ```
    FORM: X
    ```
    will be converted to 
    ```
    FORM: A
    ```

- The script looks for the key section by searching for the string "FORM X KEY" within the DOCX file. Ensure that your original document includes a key section at the end of file. The key section should provide the correct answers corresponding to each question.

    *Example*

    ```
    FORM X KEY:
        1.  A
        2.  B
        3.  C

    ```
- To help you understand the expected format of the original DOCX file, an example document (`example.docx`) has been provided in the repository. You can refer to this example to see how questions, answer choices, form numbering, and the key section should be formatted.


## Contributing

Contributions to improve the script's functionality or documentation are welcome. Please submit pull requests through GitHub.

## License

This project is licensed under the [MIT License](LICENSE).