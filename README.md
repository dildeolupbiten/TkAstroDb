# TkAstroDb

**TkAstroDb** is a Python program that uses [Astro-Databank](https://www.astro.com/astro-databank/Main_Page) to conduct statistical studies in astrology. Because of the license conditions, [Astro-Databank](https://www.astro.com/astro-databank/Main_Page)  can not be shared with third party users. Therefore those who are interested in using this program with [Astro-Databank](https://www.astro.com/astro-databank/Main_Page), should contact with the webmaster of [Astrodienst](http://www.astro.com) to get a license.

**With [Astro-Databank](https://www.astro.com/astro-databank/Main_Page)** 

If third party users have a license, they should follow the below instructions:

**1.** After got the license and downloaded the database which is in a zipped xml file from http://www.astro.com/adbexport/, in order to run the program, the xml file has to be put in the same directory with **TkAstroDb.py** script file.

**2.** Before running the program make sure that **TkAstroDb** directory tree contains at least the following:

    TkAstroDb+
             |_Eph+
             |_adb_export_181128_2309.xml
             |_TkAstroDb.py
             
**Without [Astro-Databank](https://www.astro.com/astro-databank/Main_Page)**

If third party users don't have a licence, they still can continue using the program with an empty [SQL](https://www.sqlite.org/index.html) database which users can add records in. If users have also [Astro-Databank](https://www.astro.com/astro-databank/Main_Page) xml file, both databases are merged.

**1.** Before running the program make sure that **TkAstroDb** directory tree contains at least the following:

    TkAstroDb+
             |_Eph+
             |_TkAstroDb.py

## Availability

Windows, Linux and MacOSX

## Dependencies

In order to run **TkAstroDb**, at least [Python](https://www.python.org/)'s 3.6 version must be installed on your computer. Note that in order to use [Python](https://www.python.org/) on the command prompt, [Python](https://www.python.org/) should be added to the PATH. There is no need to install manually the libraries that are used by the program. When the program first runs, the necessary libraries will be downloaded and installed automatically.

## Usage

**1.** Run the program by writing the below to **cmd** for Windows or to **bash** for Unix.

**For Unix**

    python3 TkAstroDb.py

**For Windows**

    python TkAstroDb.py

**Note:** When the program first run in Windows, users will get a [PermissionError](https://docs.python.org/3/library/exceptions.html#PermissionError)  during the installation of [Pyswisseph](https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyswisseph) library unless they run **cmd** as Administrator.

**2.** Users should see a window within one minute which is similar to below.

**Note:** Waiting time depends whether the xml file ([Astro-Databank](https://www.astro.com/astro-databank/Main_Page)) is in **TkAstroDb** folder or how many records does [SQL](https://www.sqlite.org/index.html) database include.

![img1](https://user-images.githubusercontent.com/29302909/55158048-09d9d000-516f-11e9-8c2d-3a86a1a537f7.png)

**3.** If users want to add single records to the displayed records according to the selection, they should type the name of the record. For example suppose a user wants to add **Albert Einstein** to the displayed records, the user should write **Einstein, Albert** to the **Search A Record By Name** section. While typing the name, if the record is found, the user will see an **Add** button which is used for adding records to the treeview.

![img2](https://user-images.githubusercontent.com/29302909/55158174-4dccd500-516f-11e9-8a35-a1b76224357f.png)

![img3](https://user-images.githubusercontent.com/29302909/55158243-7ce34680-516f-11e9-9446-5574ef1b9459.png)

**4.** Click **Select** button which is near to **Categories** label. After that a small window should be opened as below. Select one category or more categories to study, then press **Apply** button.

![img4](https://user-images.githubusercontent.com/29302909/50767418-d1bca280-128d-11e9-9c30-a7ce39884f35.png)

**5.** Click **Select** button which is near to **Rodden Rating** label. A small window should be opened as below. Select one rating or more ratings then press **Apply** button.

![img5](https://user-images.githubusercontent.com/29302909/50359618-72148680-056e-11e9-94e0-17938c41d268.png)

**6.** Before clicked **Display Records** button, users can click the check buttons in order to filter the records. Because some human records can be in event categories or some event records can be in human categories. But if users want to display all records, they should not click the check buttons. Then, users can click **Display Records** button. After clicked to that button, users should wait a bit. Finally the records will be displayed at treeview as follows:

![img6](https://user-images.githubusercontent.com/29302909/55159608-83bf8880-5172-11e9-8ea8-a2fecc98fcdf.png)

**7.** Users can focus on a record then by right clicking to that record they can see some options that can be done with the focused record. One of the options is deleting the record from displayed results. The deleted record will not be used in the computation later. The other option opens the [Astro-Databank](https://www.astro.com/astro-databank/Main_Page) webpage of the selected record. And the last option display the chart of the focused record.

![img7](https://user-images.githubusercontent.com/29302909/56081437-12403500-5e16-11e9-8298-702cc83d12ca.png)

![img8](https://user-images.githubusercontent.com/29302909/56081446-24ba6e80-5e16-11e9-9976-4e7418196d9d.png)

**8.** If users click **Export** menu button, they can see **Adb Links**, **Latitude Frequency** and **Year Frequency** options. By clicking **Adb Links** option, they can export the links of displayed records to [links.txt](https://www.dropbox.com/s/l69mhy5v5341gyr/links.txt) file. This file will be created at **TkAstroDb** directory. By clicking the **Latitude Frequency** option, the latitude intervals of the selected categories and the mean latitude value will be written inside [latitude-frequency.txt](https://www.dropbox.com/s/gncr0ywdk2xy056/latitude-frequency.txt) file. By clicking the **Year Frequency** option, a windows is opened as below. Users can specify the maximum, minimum and step values. After clicked the **Apply** button, the frequency of the years of displayed records will be put in a [year-frequency.txt](https://www.dropbox.com/s/2hx9gvstf0tso9r/year-frequency.txt) file.

![img9](https://user-images.githubusercontent.com/29302909/51045381-80394e00-15d4-11e9-8eed-881fa66f0afb.png)

**9.** If users click **Options** menu button, they can see **House System** option. And by clicking that menu button, they can define the house system they want to use. If they don't click this menu button, the house system will be defined according to the default setting. The default house system is defined as **Placidus**.

![img10](https://user-images.githubusercontent.com/29302909/51045506-d1e1d880-15d4-11e9-88f6-b63011273862.png)

**10.** If users click **Options** menu button, they can see **Orb Factor** option. And by clicking that menu button, they can define the orb factors for each astrological aspect. If users don't click this menu button, the orb factors will be defined according to their default settings. The default orb factors is defined as follow:

![img11](https://user-images.githubusercontent.com/29302909/50407124-2cf88a80-07e2-11e9-92f4-d51a4f7a6697.png)

**11.** If users click **Records** menu button, they can see two options which names are as **Add New Record** and **Edit & Delete Records**. If users select **Add New Record** option, they will see a new window which is as follows:

![img12](https://user-images.githubusercontent.com/29302909/55292823-eb3a3a00-53f7-11e9-8051-ffab3d7ce8ce.png)

**12.** Users should fill all the entry fields that can be seen on the image. They can select any existing categories or they can define new categories using entry field if they want. Finally users should press the **Apply** button, then the selected or defined categories will be added to the category listbox. Note that users should use decimal latitude and longitude coordinates for a place.

![img13](https://user-images.githubusercontent.com/29302909/56122350-e8545300-5f7a-11e9-846d-745359da7cfe.png)

**13.** After added new records to an alternative [SQL](https://www.sqlite.org/index.html) database which name is **TkAstroDb.db**. After added or edited or deleted the records if users click the **Select** button which is near the **Categories** label, the newly defined category can be seen.

![img14](https://user-images.githubusercontent.com/29302909/55161277-2fb6a300-5176-11e9-8d0c-db32a4ac623e.png)

**14.** After selected the newly added category and specified the **Rodden Rating**, the newly added record can be displayed.

![img15](https://user-images.githubusercontent.com/29302909/55326449-7404b480-5490-11e9-8b4b-4146dd5decb2.png)

**15.** If users click **Edit & Delete Records**, they will see a new window which is as below. Users can see all the records of the [SQL](https://www.sqlite.org/index.html) database. If users correctly typed the record name to the search entry which is under **Search A Record By Name** label, the information of the found record will be displayed in a record window. 

![img16](https://user-images.githubusercontent.com/29302909/55294306-39a40480-5409-11e9-948e-33945cd46e48.png)

**16.** If users right click after selected a record, a right click menu will occur as can be seen on following image.

![img17](https://user-images.githubusercontent.com/29302909/56081482-98f51200-5e16-11e9-82aa-58bb4b82a468.png)

**17.** If users select **Edit** option, a panel will be formed which contains the information of the selected record. After required changes have been done, users should click  **Apply** button to finish the editing process.

![img18](https://user-images.githubusercontent.com/29302909/56122350-e8545300-5f7a-11e9-846d-745359da7cfe.png)

**18.** If users select **Delete** option, the selected record is removed from the [SQL](https://www.sqlite.org/index.html) database.

**19.** If users select **Open Chart** option, the chart of the selected record will be displayed.

![img19](https://user-images.githubusercontent.com/29302909/56081607-d017f300-5e17-11e9-8cba-89e8e0d421c6.png)

**20.** If users click **Calculations** menu button, they can see options which are as follows: **Find Observed Values**, **Find Expected Values**, **Find Chi-Square Values**, **Find Effect Size Values** and **Find Cohen's D Effect Size Values**. If users want to find the astrological pattern distributions of any category, they should click **Find Observed Values** button. After clicked that menu button, a progress bar should be created as follows:

![img20](https://user-images.githubusercontent.com/29302909/55159732-cc774180-5172-11e9-88b8-14ad17ccd6c2.png)

**21.** After the computation finished, a log file (**output.log**) and an excel spreadsheet file (**observed_values.xlsx**) can be found inside nested directories like **Vocation/Occult_Fields/Astrologer/Rodden_Rating_AA/Orb_Factor_6_2_2_4_2_6_6_2_2_2_6/House_System_Placidus**. The directory names can be different according to the settings selected by the users. The spreadsheet file contains the astrological pattern distributions of displayed records.

![img21](https://user-images.githubusercontent.com/29302909/57989857-a0f54100-7aa9-11e9-9833-2060284bf812.png)

**22.** In order to calculate the expected values, the users must have two tables which include the astrological pattern distributions of two different categories. The expected values are calculated by comparing this two different categories. One category will be used as a *control group*, the other category will be used as a *research group*. While the table which is wanted to use as a *control group* should be renamed as **control_group.xlsx**, there is no need to change the name of *research group*, so it's name should be **observed_values.xlsx**. Note that users should copy the related tables to the **TkAstroDb** folder, then users can click **Calculations** menu button and they should select **Find Expected Values** option. There are two different methods to calculate the expected results.

**22.1. Flavia's Method:** When one of the categories is not a sub category of another, using this method is recommended.
    
```python3
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
       
       
def formula(x: list, y: list):
    return [sum(x) * (x[i] + y[i]) / (sum(x) + sum(y)) for i in range(12)],\
        [sum(y) * (x[i] + y[i]) / (sum(x) + sum(y)) for i in range(12)]
```
       
**22.2. Sjoerd's Method:** When one of the categories is a sub category of another, using this method is recommended.
       
```python3
#!/usr/bin/env python3
# -*- coding: utf-8 -*-


def formula(x: list, y: list):
    return [i * sum(x) / sum(y) for i in y],\
        [i * sum(y) / sum(x) for i in x] 
```

**23.** After the calculation finished, **control_group.xlsx** and **observed_values.xlsx** files will be deleted and a new excel spreadsheet file will be created as **expected_values.xlsx**. And it is recommended to move **expected_values.xlsx** file to the related category folder:

![img22](https://user-images.githubusercontent.com/29302909/57989859-a488c800-7aa9-11e9-9f3f-b4f9bf0952a3.png)

**24.** In order to calculate the chi-square values, users should copy **expected_values.xlsx** and **observed_values.xlsx** files to **TkAstroDb** folder. Then they should select **Calculations** menu button. After that, they should click **Find Chi-Square Values** option. After the compuation finished, **expected_values.xlsx** and **observed_values.xlsx** files will be deleted and a new excel spreadsheed file will be created as **chi-square.xlsx** in **TkAstroDb** folder. It is recommended that the users cut the file from **TkAstroDb** directory then paste it to the related directory. And it is recommended to move **chi-square.xlsx** file to the related folder:

![img23](https://user-images.githubusercontent.com/29302909/57989862-a783b880-7aa9-11e9-96e9-aac9a15e8c5a.png)

**25.** In order to calculate the effect size values, users should copy **expected_values.xlsx** and **observed_values.xlsx** files to the **TkAstroDb** directory. Then they should select **Calculations** menu button. After that, they should select **Find Effect Size Values** option. After the compuation finished, **expected_values.xlsx** and **observed_values.xlsx** will be deleted and a new excel spreadsheed file will be created as **effect-size.xlsx** in **TkAstroDb** folder. It is recommended that the users cut this file from **TkAstroDb** directory then paste it to the related directory. And it is recommended to move **effect-size.xlsx** file to the related folder:

![img24](https://user-images.githubusercontent.com/29302909/57989863-aa7ea900-7aa9-11e9-8ae0-be1b525939fe.png)

**26.** In order to calculate the Cohen's d effect size values, users should copy **expected_values.xlsx** and **observed_values.xlsx** files to the **TkAstroDb** directory. Then they should select **Calculations** menu button. After that, they should select **Find Cohen's D Effect Size Values** option. After the compuation finished, **expected_values.xlsx** and **observed_values.xlsx** will be deleted and a new excel spreadsheed file will be created as **cohens_d_effect.xlsx** in **TkAstroDb** folder. It is recommended that the users cut this file from **TkAstroDb** directory then paste it to the related directory. And it is recommended to move **cohens_d_effect.xlsx** file to the related folder:

![img25](https://user-images.githubusercontent.com/29302909/57989865-ace10300-7aa9-11e9-8c07-697cfd2346fe.png)

## Spreadsheets

[observed_values.ods](https://www.dropbox.com/s/9au6e09qb65c3em/observed_values.ods)

![img26](https://user-images.githubusercontent.com/29302909/55176661-9944aa80-5192-11e9-8891-ca4498541e8b.png)

[expected_values.ods](https://www.dropbox.com/s/5sm9jbpuupeh45o/expected_values.ods)

![img27](https://user-images.githubusercontent.com/29302909/55176770-d27d1a80-5192-11e9-95e9-9aee8916ccfb.png)

[chi-square.ods](https://www.dropbox.com/s/myq4sws3kqk3jbl/chi-square.ods)

![img28](https://user-images.githubusercontent.com/29302909/55176879-0c4e2100-5193-11e9-923f-dfe0c3b4a934.png)

[effect-size.ods](https://www.dropbox.com/s/gnppq4cym6g3cx2/effect-size.ods)

![img29](https://user-images.githubusercontent.com/29302909/55176960-369fde80-5193-11e9-979a-f7429821261d.png)

[cohens_d_effect.ods](https://www.dropbox.com/s/bu510gdehi9md5e/cohens_d_effect.ods)

![img30](https://user-images.githubusercontent.com/29302909/58037785-e0b63a00-7b36-11e9-9868-89b6154d3717.png)

## Notes

**1.** The tables may not be opened by Microsoft Excel. Therefore it is recommended to use [Libre Office](https://www.libreoffice.org/download/download/). 

**2.** If users want to put their files in a cloud system like Dropbox, it is recommended that the format of the excel files should be changed from *xlsx* to *ods* format.

## Licenses

TkAstroDb is released under the terms of the GNU GENERAL PUBLIC LICENSE. Please refer to the LICENSE file.
