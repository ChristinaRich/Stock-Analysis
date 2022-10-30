# Stock-Analysis with VBA

## Project Overview

The purpose of this project was to analyze daily volume and returns for 12 different stocks for 2017 and 2018. The original analysis we performed was refactored in order to improve efficiency in the model and reduce time taken to process the data. Using VBA, an output page was created that the user can generate through entering the year they'd like to analyze. 


## Results

Refactoring the original code decreased the time it took to process the code. My first code needed almost two seconds to process. The new code was a little more than half of a second. On a larger dataset, this difference would be much more meaningful. 

I think that eliminating the nested loop saved some time. The code below is the original code. With the nested loop, our variables switched from i to j. The second loop is running within the first loop, which isn't necessary, causing more iterations than necessary.

![ScreenShot](https://github.com/ChristinaRich/Stock-Analysis/blob/c12a329e1f27ceca9eae0303c56a5430525804b0/Original%20Code.png)



This second screenshot below closes the first loop, and then creates a second loop that is outside of the first. I believe this created the biggest boost to efficiency.


![ScreenShot](https://github.com/ChristinaRich/Stock-Analysis/blob/c12a329e1f27ceca9eae0303c56a5430525804b0/Refactored.png)


## Summary

### Advantages and Disadvantages of Refactoring

Refactoring makes our code more efficient. In this refactoring analysis, we simplified the code using fewer steps and clearer code. This not only uses less memory (decreasing the time to process the macro), it also helps other users to read clearer code. Refactoring takes more time to do up front because you have to make changes and analyze opportunities for efficiency, but it is worth the effort for the efficiency savings later. As we saw with this example, the refactoring of the original VBA script did exactly this. We decreased the runtime, and it is much easier to follow going forward for other users. 




