# Stock Decisions with VBA

## Overview of Project
This project takes an array of data collected following 12 stocks and utilizes 12 stocks and utilizes VBA code to quickly and automatically gain insights towards the performance of each of the stocks. The output gives an easily readable worksheet equipped with buttons that allow the code to be run.

### Purpose
The purpose of this project was to follow up on several investment stocks and track their returns. Utilizing several sets of macros that combed through the full range of data the returns of each stock was determined on a yearly basis. This project also demonstrates how VBA can be used effectively for data analysis, as well as how reformatting a code can drastically affect the speed at which the code runs.

## Results
![VBA_Challenge_2017](https://user-images.githubusercontent.com/103979048/172708212-3ffdfa89-4a27-4101-bd12-f8df701db84d.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/103979048/172708241-6752f637-9abd-4ae1-b36b-655e8101a4ec.png)


### Stock returns
Almost all the stocks followed saw a good degree of return in the year 2017. The one stock that had a negative return was TERP. However, in 2018 many of the stocks turned and had negative returns. ENPH and RUN were two stocks that had 2 consecutive years of positive return, whereas stocks such as DQ, JKS, and SPWR went from healthy returns to pretty large losses, the largest swing being DQ.

###Program Efficiency
![Old2017](https://user-images.githubusercontent.com/103979048/172708311-ebdbf528-6527-4a0d-829e-b0e6311bae39.png)
![Old2018](https://user-images.githubusercontent.com/103979048/172708327-e90c1973-67d5-4851-a5af-a87a04d27269.png)


Reworking the vba code drastically increased the speed that the macro ran. This is evident in the amount of time each macro took to run. The reworked code for 2018 saw a 87 % decrease in run time or in other words it ran 8.2 times faster. One key difference in the code was the use of the nested for loops. In the original code, the code would go through the full length of the data twelve separate times because of the nest seen here.
![OldForLoop](https://user-images.githubusercontent.com/103979048/172708379-cfe4da12-7f73-4556-b4dc-0193706909a0.png)

 With the new code the rows are searched through once
 ![CHallengeForLoop](https://user-images.githubusercontent.com/103979048/172708439-9c110282-90e7-46a3-bde5-135105e434c7.png)

and the stock index changes as the code reaches a new stock.
![ChallengeChangeStock](https://user-images.githubusercontent.com/103979048/172708502-8144f9ed-2f06-4bec-bddb-d9a8f0f8c194.png)

 Another set of differences that likely decreased the run time was the creation and initialization of arrays for each of the ticker variables.

##Summary
###General code refactoring
Code refactoring can be immensely beneficial. Reducing the run time of a program can be the difference between success and failure particularly when the code will be applied over significant amounts of data. This also allows the processors to put in less work and for less energy to be consumed. There are some disadvantages to refactoring code, however. First, refactoring code can have unintended consequences that lead to bugs. If one is not careful on how they implement the refactor the process of debugging could end up undoing the whole code. Often, it will be difficult to be able to see what needs to be refactored particularly if you were the one working on the code. In this case it would be beneficial to have others look over the code but this means additional resources devoted to this particular code. Finally, and least importantly, sometimes the most efficient solution might come at the cost of the elegance of the code. In any case, many of these disadvantages can be combated by a healthy amount of comments and pseudocode so that each block of code’s purpose is clear.

###Challenge code refactoring
The code ran significantly faster after refactoring. Additionally it was convenient that the reformatting was included into the macro instead of having separate macros for reformatting and calculating. In my implementation of the refactor I ran into several syntax bugs that lengthened the process of writing after the initial refactor was complete. Luckily most bugs were clear to see and were easy enough to dispatch.
