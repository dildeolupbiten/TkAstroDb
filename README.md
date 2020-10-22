# TkAstroDb

**TkAstroDb** is a Python program that uses [Astro-Databank](https://www.astro.com/astro-databank/Main_Page) to conduct statistical studies in astrology. Because of the license conditions, [Astro-Databank](https://www.astro.com/astro-databank/Main_Page) can not be shared with third party users. Therefore those who are interested in using this program with [Astro-Databank](https://www.astro.com/astro-databank/Main_Page), should contact with the webmaster of [Astrodienst](http://www.astro.com) to get a license.

After downloading the program, users should see the below files and folders in the main directory of the program.

![img1](https://user-images.githubusercontent.com/29302909/96886176-006c3300-148c-11eb-8a05-24e70029a08a.png)

## Availability

Windows, Linux and macOS

## Dependencies

In order to run **TkAstroDb**, at least [Python](https://www.python.org/)'s 3.6 version must be installed on your computer. Note that in order to use [Python](https://www.python.org/) on the command prompt, [Python](https://www.python.org/) should be added to the PATH. There is no need to install manually the libraries that are used by the program. When the program first runs, the necessary libraries will be downloaded and installed automatically.

## Usage

**1.** Run the program by writing the below to **cmd** for Windows or to **bash** for Unix.

**For Unix**

    python3 run.py

**For Windows**

    python run.py
    
**Note:** In order to run the program, Windows users could double click the `run.bat` file.

**Note:** When the program first run in Windows, users might get a [PermissionError](https://docs.python.org/3/library/exceptions.html#PermissionError)  during the installation of [Pyswisseph](https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyswisseph) library unless they run **cmd** as Administrator.

**2.** Short time later users should see a window which is similar to below. Waiting time depends on the properties of the user's computer.

![img2](https://user-images.githubusercontent.com/29302909/96886659-77a1c700-148c-11eb-86fb-036db92fdae4.png)

**3.** By default there's no folder called **Database**. If users click the **Database** menu button, an empty folder called **Database** would be created in the main directory of the program. Users should move the **XML** file that they obtained to this folder. And if the users click the **Database** menu button for the second time, a window as below would occur.

![img3](https://user-images.githubusercontent.com/29302909/96887895-b2f0c580-148d-11eb-91ae-d730a4735fe1.png)

Users can move many databases to the **Database** folder and select the database you want to work with. The above window is only destroyed when the users press the **Apply** button.

**Note:** Users can also use a database that is derived from the original [Astro-Databank](https://www.astro.com/astro-databank/Main_Page) via using [TkEnneagram](https://github.com/dildeolupbiten/TkEnneagram). The format of the derived database is JSON. What users should do is that they should move the derived database to the **Database** folder.

**4.** A new frame should cover the main window as below in a few seconds after pressed the **Apply** button.

![img4](https://user-images.githubusercontent.com/29302909/96889338-2d6e1500-148f-11eb-98f1-3d100c38ae25.png) 

**3.** If users want to add records manually to the displayed records, they should type the name of the record in the combobox. For example suppose a user wants to add **Albert Einstein** to the displayed records, the user could write **einstein** or a keyword that the program could find to the **Search A Record By Name** section, then if users press **Enter** key, a list of records that contains **einstein** characters would be inserted to the combobox and a drop-down menu would be popped. Users can select the found records from this drop-down menu via clicking to the arrow of the combobox. After selecting the record, a new button called **Add** is created as below.

![img5](https://user-images.githubusercontent.com/29302909/96890707-93a76780-1490-11eb-957b-071d69a10a97.png)

If users click the **Add** button, the record would be added to the treeview and the **Add** button would be destroyed. However, users can continue selecting the records from the drop-down menu. The **Add** button would be created for the next record unless the record is already in the treeview.

![img6](https://user-images.githubusercontent.com/29302909/96890745-9efa9300-1490-11eb-8455-6034c44e4c57.png)

**4.** If a record inserted to treeview is selected and users use the right-click of their mause, a right click menu would open and if they want, users could open the ADB page of the record or remove the record from the treeview.

**5.** Users could select the categories of ADB via clicking the **Select** button near the **Categories** label. There are two ways of selecting the categories. One is the **Basic** category selection method which is coming by default and the window as below would open if the selection method is **Basic**.

![img7](https://user-images.githubusercontent.com/29302909/96892541-68257c80-1492-11eb-904e-f79b9cc9bf21.png)

Users can search a category by writing something to the search entry. If the characters that users typed match with the characters of the categories in the category list, the horizontal scrollbar would move to the category that contains the characters. And if users press **Enter** key, the program would move to the other categories that contain the characters.

In order to select a category, users should use the right-click of the mause and select the **Add** option. The color of the added category would turn to red. Users could select all the categories by using `CTRL-A`, then the selected categories should be added.

In order to apply the selections, users should click the **Apply** button.

The other category selection method is called as **Advanced**. Users could change the category selection method via coming to the **Options** menu cascade and clicking the **Category Selection** menu. If the **Category Selection** menu button is clicked, a window as below would open.

![img8](https://user-images.githubusercontent.com/29302909/96896860-b0df3480-1496-11eb-86eb-f9ce3c3df6f5.png)

As mentioned before, by default, the **Basic** option is selected. The selected option would be valid for the next time.

If users click the **Select** button near the **Categories** after selected the **Advanced** category selection, a window as below would open.

![img9](https://user-images.githubusercontent.com/29302909/96897416-3cf15c00-1497-11eb-96f9-7189e6ba6f09.png)

The left frame is for including the categories whereas the right frame for ignoring the categories. For example suppose a user want to create a control group for **Cancer** category that include **Non-Cancer** records, the user could select **Cancer** category from the right frame to be ignored. Note that all selected categories should be added before pressing the **Apply** button.

**6.** Users could select the [Rodden Ratings](https://www.google.com/search?&q=rodden+rating) they want to include, thus only the records that have the selected Rodden Ratings would be inserted to the treeview. If users click the **Select** button near the **Rodden Rating*, a window as below would open.

![img10](https://user-images.githubusercontent.com/29302909/96898559-88583a00-1498-11eb-948e-6327ea3905f9.png)

**7.** After selected the **Categories** and **Rodden Ratings**, users should click the **Display** button to insert the filtered records to the treeview. The inserting process may take time depending on the selected categories and rodden ratings. After the inserting process is completed, users should receive an information message as below.

![img11](https://user-images.githubusercontent.com/29302909/96899042-146a6180-1499-11eb-8e7e-8236f53cc4ea.png)

Users could also filter the records using the checkbuttons which can be seen on the main window.

**8.** If users click the **Export** menu cascade after the records are inserted to the treeview, users can export the links, the latitude and year frequencies of the records. The files would be created in the main directory. And it's recommended to move these files to the folder that would be created after completed the process of finding the observed values.

**9.** Before passing to the calculation process, users could select the **House System** to be used by coming to the **Options** menu cascade then clicking the **House System** menu button. 

![img12](https://user-images.githubusercontent.com/29302909/96900253-91e2a180-149a-11eb-94cb-ca6d7d7f57c2.png)

By default the **Placidus** house system is selected. However the selected house system would be used as default for next calculations.

**10.** Users could also change the default **Orb Factors** by coming to the **Options** menu cascade then clicking the **Orb Factors**.

![img13](https://user-images.githubusercontent.com/29302909/96900630-fd2c7380-149a-11eb-826c-0d1cfc60b9fb.png)

The default values of the **Orb Factors** can be seen above. However if users change the values, the changed values would be used as default for next calculations.

**11.** When users completed selecting the options they want, they can start the process of finding the observed values. In order to do that, users should come to the **Calculations** menu cascade then click the **Find Observed Values** menu button. Immediately, the calculation would start and users could watch the progress at the main window.

![img14](https://user-images.githubusercontent.com/29302909/96901528-108c0e80-149c-11eb-99ca-a8caa9c64130.png)

After the calculation is completed, users should see a message that states that the calculation is completed. And also users should see a new nested directory in the main directory that includes the **output.log** and the **observed_values.xlsx** files. The name of the directories would change according to the selected category, house system, rodden rating, checkbuttons.

![img15](https://user-images.githubusercontent.com/29302909/96901522-0f5ae180-149c-11eb-96e9-f293e1910a0e.png)

The **output.log** files includes the information of process. If an error occurs during the calculation process, the errors would be written to the **output.log** file. 
 
**12.** Users could use other menu buttons of the **Calculations** menu cascade after they found the observed values of different categories. Users could create control groups by selecting all the categories in the category list. That control group would be the largest control group that represents the whole database. However users could create even smaller control groups.

A control group could be both an indepenent category or a super category. Regarding the kind of the control group, users should specify the method that will be used in calculating the expected values.

For instance, if the user wants to use a super category as the control group, the method should be selected as **Subcategory**. Already by default the method is selected as **Subcategory**. However if users want to change the method of calculating the expected values, they should come to the **Options** menu cascade and select the **Method** menu button. After clicked this menu button, a window like below should open.

![img16](https://user-images.githubusercontent.com/29302909/96903276-4d590500-149e-11eb-880a-0e315669a86a.png)

So, basically the **Subcategory** method is for the cases that one category is the sub category of another and the **Independent** method is for the cases that the categories are independent.

For instance, as mentioned before, if a user wants to create a control group for **Cancer** category that includes people that have no cancer, the method should be selected as **Independent**.

In order to use the the other calculations in the **Calculations** menu cascade, the spreadsheet files should be moved to the main directory of the program and the name of the spreadsheet file that will be used as the control group should be renamed as **control_group.xlsx**. Already the program would look whether or not the files are located in the main directory and would raise warning messages if the files don't exist. If the files are located in the main directory, immediately the calculation process would be started and the process would be completed in seconds and the new spreadsheet files would be created in the main directory. It's recommended to move those files to the directory where the **observed_values.xlsx** is located in. So suppose a user created the spreadsheet files of expected values, chi square values, effect size values, cohen's d effect values and binomial limit values for **Cancer** category, it's recommended to move the created files near the **observed_values.xlsx** of the **Cancer** category.

**13.** If an update is released, users can update their program by coming to the **Help** menu cascade and clicking the **Check for updates** menu button.

**14.** If users click the **About** section which is under the **Help** menu cascade, a window like below opens and users can contact the developer using the email link written on the frame.

![img17](https://user-images.githubusercontent.com/29302909/96908127-e12dcf80-14a4-11eb-8994-7be585753cb2.png) 

## Spreadsheets

[observed_values.ods](https://www.dropbox.com/s/95hrm53bmy8pqxt/observed_values.ods)

![observed values](https://user-images.githubusercontent.com/29302909/96907424-dfafd780-14a3-11eb-9714-52846925653d.jpeg)

[expected_values.ods](https://www.dropbox.com/s/9sx39qm7jj9u76q/expected_values.ods)

![expected values](https://user-images.githubusercontent.com/29302909/96907420-dde61400-14a3-11eb-8b81-6e8bbd5366ff.jpeg)

[chi-square.ods](https://www.dropbox.com/s/frviu3e9y0cjss2/chi-square.ods)

![chi-squre values](https://user-images.githubusercontent.com/29302909/96907407-da528d00-14a3-11eb-9cef-527080eb552a.jpeg)

[effect-size.ods](https://www.dropbox.com/s/wdt4tu34d758gta/effect-size.ods)

![effect size values](https://user-images.githubusercontent.com/29302909/96907412-dc1c5080-14a3-11eb-96f1-f585723ef84d.jpeg)

[cohens_d_effect.ods](https://www.dropbox.com/s/k4c37cm43w0jjdh/cohens_d_effect.ods)

![cohen's d effect values](https://user-images.githubusercontent.com/29302909/96907396-d888c980-14a3-11eb-8bab-ce080b1df81c.jpeg)

[binomial_limit.ods](https://www.dropbox.com/s/kco8ssmi78344nd/binomial_limit.ods)

![binomial limit values](https://user-images.githubusercontent.com/29302909/96907367-d161bb80-14a3-11eb-9c91-d5d0804cd669.jpeg)

## Notes

**1.** The tables may not be opened by Microsoft Excel. Therefore it is recommended to use [Libre Office](https://www.libreoffice.org/download/download/). 

**2.** If users want to put their files in a cloud system like Dropbox, it is recommended that the format of the excel files should be changed from *xlsx* to *ods* format.

## Licenses

TkAstroDb is released under the terms of the GNU GENERAL PUBLIC LICENSE. Please refer to the LICENSE file.
