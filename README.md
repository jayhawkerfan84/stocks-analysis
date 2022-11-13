# VBA Challenge

## Overview of Project

I have a friend Steve that recently graduated from college and is looking to help his 

parents diversify their investment portfolio. Steve’s parents have invested all their 

money into New Energy Corporation, DAQO, but Steve has come to me to investigate other 

stocks worth investing in. Steve has provided me a spreadsheet filled with green stocks to 

analyze and give feedback on a company to invest in. 



## Results

To analyze the stocks Steve provided we needed to build a VBA script that will loop through the 

stocks total volume and percent of return in 2018. We also needed to format the spreadsheet so the 

results would be easy for Steve to read. Our VBA code was written in such a way that Steve could 

continue to reuse the code against other data sets in the same format. Meaning, Steven provided 

just years 2018 and 2017. But he can use the code to analyze stocks moving forward, 2019, 2020, 

etc. We also added buttons so Steve and or Steve’s parents do not need to understand the code but 

just to click on the buttons to run the code to see the results. 



![2018 Green Stock Analysis](https://github.com/jayhawkerfan84/stocks-analysis/blob/main/resources/VBA_Challenge_2018%20(green_stock).png)

The results were pretty clear as to the best stocks to invest in, ENPH and RUN. Both of these 

stocks had greater than 80% return in 2018. 

![Refactored Code](https://github.com/jayhawkerfan84/stocks-analysis/blob/main/resources/VBA_Challenge_2018.png)

We later refactored our code so it was more efficient. While the output is the same, you can see 

the time to run the code compare to the screen shot above is faster. This is helpful if Steve 

needed to run the script against a larger data set. 

## Summary

### Reasons for Refactoring

  There could be many reasons for refactoring. Maybe the code base you are working on is old and 

built in a monolith style and you would like to break it up into micro services. Maybe pieces of 

the code have “quick fixes” due to production issues and you need to go back through and write 

better code. I feel that tech departments should always be looking to refactor code. If you think 

about it, a tech department could be deploying as frequent as every day. The more and more code you 

create the harder it will be to manage the code. Looking for areas to refactor will continue to 

make the code easier to work with. An analogy one could use it think about refactoring as taking 

your car to the shop for the an oil change and general maintenance. The longer you go without doing 

this the higher risk there is to your car breaking down. 

### Reasons Against Refactoring
	

  The only challenges I see with refactoring would be lack of domain knowledge. Maybe the code 
  
you’re looking to refactor is very old and the developer that wrote the code no longer works for 
  
the company. Or maybe the code is written in a language that the person refactoring the code is 
  
not familiar with. From a business perspective, operations usually wants their needs met and 
  
refactoring doesn’t move the needle for them. Therefore they might ask to hold off on 
  
refactoring. 

### Advantages/Disadvantages to the VBA script

  I think the advantages to the refactored VBA code would be the code runs more efficiently than 
  
the original script. We had already removed the hard coding in the original script, which would 

have been something to remove if it was still in place.  The disadvantage to the refactored script 

was the lack of experience writing such a script. It took me a while to figure out the arrays the 

challenge was asking me to add to the script. 

