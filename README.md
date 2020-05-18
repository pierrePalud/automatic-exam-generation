# automatic-exam-generation

Automatic exam generation from pool of questions.


## How to use ?

Simply double click on the `exam-random-creation.exe` file, and answer a few questions.

## What will the program do ?

It will import the whole pool of questions you created, and use it to
1. create as many students exams as you wish (in the `students_exam` folder) and the corresponding corrections (in the `corrections` folder)
2. create a new excel file, which is basically a copy of your `questions_pool.xlsx` with extra columns. Each column correponds to a test version. These new columns shows if a question has been selected to the nth version or not, and if so, at which number.