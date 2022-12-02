# Finance Forecasting - Monte Carlo Simulations
Welcome to the Finance Forecasting Monte Carlo Simulation project. The program uses a database “MonteCarlo” with values corresponding to revenues, costs, and standard deviations. The database consists of 3 options/forecasting types:
- Balance (balanced revenues and expenses)
- Low (low revenues, high expenses)
- High (high revenues, low expenses)
- 
![image](https://user-images.githubusercontent.com/116671665/205357570-6a1e7d5d-2239-4a62-ae57-b6e4de83408e.png)
![image](https://user-images.githubusercontent.com/116671665/205357600-e32a5c44-efd1-468c-9b86-9483f2477269.png)

The values corresponding to the strategy chosen by the user are imported into the excel file and displayed

![image](https://user-images.githubusercontent.com/116671665/205357673-99c4085e-05a8-4540-9d0e-253514026153.png)

With these values now on the excel file the first simulation is run using normal distribution using parameters with a random value, and the standard deviation of revenues and variable expenses. 

![image](https://user-images.githubusercontent.com/116671665/205357707-823f8633-2e12-40f3-9ba2-f299cdc9d194.png)

Below you will see the user form created to interact with the user, retrieve their selections and inputs to perform the program.
Typically, with Monte Carlo simulations, the more simulations the more accurate the result of the prediction. Although,  because the program must loop through and analyze each result, this can take a while to finish compiling

![image](https://user-images.githubusercontent.com/116671665/205357789-ab7b7d26-8c13-423b-988f-4a3f74a2d686.png)

As you can see to the right, each iteration is outputted to the excel file assigning each iteration a number. Any value that is greater than 0 is displayed in green (net profit), and any number less than 0 is displayed in red (net loss). 

![image](https://user-images.githubusercontent.com/116671665/205357855-e313aef4-4d42-4164-b781-97190d4cc060.png)

Once the simulations have been completed, the program loops through each iteration and calculates the percentage of iterations that resulted in a net loss and a percentage of the whole. It also calculates the percentage of iterations that resulted in a profit greater than the profit goal inputted by the user. The program displays a description of the strategy they’ve chosen in the form of a message box and a description of the program. They will also see a message appear summarizing their chance of losing money, and the chance they hit their profit goal.

![image](https://user-images.githubusercontent.com/116671665/205357906-35872807-9134-433a-b728-bf7d2092a949.png)
![image](https://user-images.githubusercontent.com/116671665/205357919-a7be5cfa-7b29-4fe1-93ae-e5a09c19f97d.png)


In the center of the excel sheet, the program displays a lined scatter plot, demonstration the movements of each consecutive iteration. Based on the returned valued, the scatter plot will display the distribution of positive and negative return values, automatically adjusting for the type of forecasting chosen by the user. This allows the user to easily compare the results of each respective forecasting option. 

![image](https://user-images.githubusercontent.com/116671665/205357952-bab4358b-9c6b-443e-9c20-3d84b035d0e7.png)

After the program is complete and the scatter plot has been displayed, the program creates a new Word document that explains the results of the program to the user. The program uses the TypeText function to type the body of the paragraph through VBA in excel. It also copies the lined scatter plot and pastes it onto the new word document for the user to see. Below is a screenshot of what the word document looks like. 

![image](https://user-images.githubusercontent.com/116671665/205357997-8293e1d1-7f71-471d-be76-77e22566fc91.png)
