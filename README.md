The Project
WORTHY is a 6 piece AI model that runs on local hardware using OpenAI API as the response mechanism.
Each aspect accomplishes soemthing not yet available to the general public. For example, the writing "W" engine can write an entire book from start to finish automatically.
My favorite use for ChatGPT so far has been Deep Research, and thus, I focused first on building the "R" research engine. It outperforms Deep Research, functions with less effort, but doubles/triples wait-time.
This model was inspired by a previous project "SheetAI" which is where I realized that spreadsheets are actually decent candidates for housing code and have some natural user-friendly advantages.
Each of the 6 engines will output into the appropriately named XLSX document.
The R engine is currently at and will outpace the effectiveness of Deep Research as it has creative iteration methods. The primary way to iterate from a high volume of responses is to evaluate them.
In order to evaluate effectively, grading becomes the primary focus on precision work, whereas the prompt must be a genius to create one. 
Other ways to improve mostly include speed like future goals to prune some areas dynanmically to reduce the search tree, though this is a complex addition.
Other ways to improve include adapting the ABC columns in Sheet 2.
Overall, R is quite good already, though has the ability to be excellent.

Currently, R and W work as of 12/4/2025
W will run until a work of writing is complete. It is designed to facilitate writing in a way that seems progressive. Therefore, it allocates chapters, sections within chapters, and everything flows together.
R will conduct research and evaluate it. It uses a GPA scale to predict how strong its research is after each successive series. 

Running these
You will need Python, Powershell, Gitbash, and Chocolatey. You will have to set up a virtual environment to support ChatGPT, which I believe can be called as pip install openai
I run commands through bash. To do so, navigate to your virtual environment, then set your own ChatGPT API key (create in OpenAI Platforms). To start, see note below
You will need an XLSX document as well as the current python version. name the python document Worthy.py, and Worthy.XLSX, similarily. 
For the XLSX, you must create 5 sheets, all labeled the stock Sheet1, Sheet2,... 
In Sheet1 cell A1, you need a dropdown menu with "W, O, R, T, H, Y", for 6 total entries.
For cell A2, you can paste the text: Enter all background info starting in A3, no gaps, only column A.
Cells A3 and below in the column must be sqeuential (no blanks). It will read anything before the first blank as a prompt. It will not read anything beyond that, or in any column except A.
1. cd ~/Downloads
2. source env/Scripts/activate
3. export OPENAI_API_KEY="x"
4. python Worthy.py
   or
4. python Worthy.py --untilgpa #.#
(for R model)

Versions
I am currently working on different versions independently. I won't upload each version but I'll mention differences between those posted.
The first 3000 lines took about 15 days or about 200 lines per day. I project that after around 10,000 lines, all of the WORTHY features will be interesting, and 15,000 they will be extremely useful.
My basis is that for R, it was just running after 800-1000 lines, and at 1250, it's pretty solid with only improvements/optimizations needed (no structural issues). 

Notes on R
With 4.1-mini, you can complete 1 series in 3 minutes. With 4.1 its closer to 6 minutes. 5.1 takes around 36 minutes.
How long and best model to optimize output? After 8-10 series, both quality and diminishing returns equalize with time. Therefore, 10 series is ideal for each model.
You can run 4.1 mini quality in 30 minutes for 10 series, whereas 5.1 would be 3 hours.
You can breakdown some of these time barriers for choosing quality thanks to department divergence. There is 1. Employee, 2. Evaluator, 3. Executive.
The bulk of work is through employee and evaluator. If you just change executive to 5.1, it is the least time loss and highest quality gained.

Notes on W
The W model is capable of writing anything effectively in spreadhseet form. After which, simply copy paste the text into a document and it will flow perfectly.
The Ai was able to produce enough subsequent text to fill 270 pages in under 1 hour. In the current version, the writing is quite good, and aware of itself and other writing in the project (cohesive).
Keep in mind, it will generally auto-select the writing type. You cant manually trigger the writing type yet. (Novel, short story, etc.) Only auto detect is currently live (Backwards I know).
