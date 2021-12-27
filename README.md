# stock-analysis

Using VBA to analyze stocks

1. Overview of Project: Explain the purpose of this analysis.

The purpose of this analysis is to make the code we used from module two able to analyze the entire stock market over 2017 and 2018. The goal is to make our script from module two run faster so Steve can analyze more stocks more efficiently. Our script will run faster by refactoring our original code, looping through all the data at once to get the same result we achieved in module 2. 

2. Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

I created the three output arrays of tickerVolumes, tickerStartingPrices, and tickerEndingPrices using the following code: 

<img width="307" alt="Screen Shot 2021-12-26 at 10 05 59 PM" src="https://user-images.githubusercontent.com/96043107/147436499-7cf59bb4-e77d-431d-9c60-1bee33946f8e.png">


This challange went relatively smoothly, but for some reason I keep encountering a debugging error when I try and run my code for 2017. The code works perfectly fine for 2018, but when I run it for 2017, I get an error code that reads: "Run-time error '13': Type mismatch". It points to the following line of code as the source of the problem: 
   <img width="382" alt="Screen Shot 2021-12-26 at 9 28 43 PM" src="https://user-images.githubusercontent.com/96043107/147435193-939b2e37-98d6-40b0-9e92-29f84fc63bde.png">
   
I am unable to determine why the code runs perfectly for year 2018, resulting in the following pop up image with a run time of 0.0703125 seconds: 
<img width="260" alt="Screen Shot 2021-12-26 at 9 39 32 PM" src="https://user-images.githubusercontent.com/96043107/147435250-278e0ac6-0fba-4200-95a4-0b7d7b4d2dec.png">

But this is my first time using VBA, and I'm pleased that I was able to get it to work for at least one of the years.. 

Compared to the original run time of 2018 within the module, our original code resulted in a run time of: 0.2578125 seconds, as seen here: <img width="255" alt="Screen Shot 2021-12-26 at 9 48 34 PM" src="https://user-images.githubusercontent.com/96043107/147435531-432d71ea-3f68-4e71-9ac4-d9506899233b.png">
So, the new code is about 70% faster than the old code from module 2, which is quite an improvement. 

Lastly, I looped through my 3 different arrays to output the ticker, the total daily volume, and the return with the following code: 

<img width="466" alt="Screen Shot 2021-12-26 at 10 06 14 PM" src="https://user-images.githubusercontent.com/96043107/147436588-7ad78a73-e9d3-426a-93d0-37dbb21bde82.png">



3. Summary:

3a: The advantages and disadvantages of refactoring code

   The main advantages of refactoring code include making the code easier for someone to follow along as to what were doing, and helps with code organization. This     would allow for overall faster coding, programming, and would help in the debugging stage. The code becomes simpler to read and understand because the data is       looped all together instead of in pieces. Sometimes, refactoring code isn't possible because of the size of some particular datasets, or in some cases if you       are experimenting with a new set of code that has not been refactored before. Overall, refactoring is preferred in order to keep code neat, easy to follow            along, and concise. 
  
3b: Summary Cont'd

How do these pros and cons apply to refactoring the original VBA script?

   The largest advantage I noticed when refactoring the original VBA script created in module 2 was the 70% increase in speed that the code was able to run in. A       faster working code would allow for more efficient stock-analysis, especially on a larger scale like Steve is planning on using our Excel file for!
