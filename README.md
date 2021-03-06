# Digital Monopoly in Excel
<p align="center">
<img src="/images/playerinfo.PNG">
</p>

<p align="center">
<img src="/images/propertycard.PNG" width=500> 
</p>

<p align="center">
<img src="/images/payments.PNG">
</p>

## Description
This is an excel file for tracking most of the components of the board game Monopoly. It displays the name of each player along with how much money they have and all of their purchased properties. Currently, the project is unfinished because Mediterranean Ave is the only purchasable property.

## Walkthrough
Once the file is open, start by clicking the **New Game** button. You will be prompted to enter each player's name. Names will automatically be filled out in the sheet, bank accounts will be reset to 1500, Free Parking will be reset to 0, and all owned properties will be cleared.

Next, click on the **Make a Transfer** button. You will be taken to the Transfers sheet. 

Click the **Pay Players** button. You will be prompted to enter the name of the paying and recieving players (use exact name spelling, case sensitive). Enter any positive integer from 1 to 1500. 

To verify the transaction, click the **Return to Player Menu** button and you will be taken to the Player Info sheet. You should see that the two players' account totals changed by the amount you entered.

Next, click the **Make a Purchase** button. Look at the first property at the top, which is Mediterranean Avenue. Click **Purchase Property** directly below the property card. You will be prompted to enter a player's name.

To verify the transaction, click the **Return to Player Menu** button in the bottom right and you will be taken back to the Player Info sheet. You should see "Mediter. Ave" in that player's properties owned, and their account total reduced by 60.

You can also pay into or collect from Free Parking.

## See the Code
To view the VBA code, go to the Developer tab at the top, and click on View Code in the Controls section.
If you do not see the Developer tab, you need to enable it. You can find instructions [here](https://www.excel-easy.com/examples/developer-tab.html).

You should see a project tree on the left. Look under VBAProject -> Microsoft Excel Objects -> Sheet 1, 2, and 3 for the code for all the buttons on their respective sheet.
Look under VBAProject -> Modules -> Module1 for all the functions called to in the code mentioned above.

## History
This was my first VBA project. I made it in April 2019 when I was learning VBA over spring break. I made some minor changes in September 2020 before uploading it here.

## Challenges
The most challenging part of this project for me was the learning curve for VBA. I had a hard time finding resources for this language compared to others like C++. Eventually I found [this website](https://www.excel-easy.com/vba.html) which provided good documentation and examples.

## Future Additions
- [ ] Better appearance
- [ ] Purchase buttons for all other properties
- [ ] Dice rolling and a visualization of the board

Contributions are welcome!
