# automatic-exam-generation

Automatic exam generation from pool of questions.

## How to install ?

Click on the green the Clone or download icon, and Dowload as zip. Then extract the content of the zip file anywhere on your computer. It is now ready to be used !

## How to use ?

First, Create a `questions_pool.xlsx` file, with the exact same structure (i.e. column names and number) as the example I created.
Then simply double click on the `exam-random-creation.exe` file, and answer a few questions to design your exam.

## What will the program do ?

It will import the whole pool of questions you created, and use it and your requests to
1. create as many students exams as you wish (word .docx files in the `students_exam` folder) and the corresponding corrections (word .docx files in the `corrections` folder)
2. create a new excel file (`questions/whole_exam_correction.xlsx`), which is basically a copy of your `questions_pool.xlsx` with extra columns. Each column correponds to a test version. These new columns shows if a question has been selected to the nth version or not, and if so, at which number.
