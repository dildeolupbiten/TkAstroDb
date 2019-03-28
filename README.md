# TkAstroDb

TkAstroDb is a program that uses [Astro-Databank](https://www.astro.com/astro-databank/Main_Page) to conduct statistical studies in astrology. Because of the license conditions, [Astro-Databank](https://www.astro.com/astro-databank/Main_Page)  can not be shared with third party users. Therefore those who are interested in using that program should contact with the webmaster of [Astrodienst](http://www.astro.com) to get a license.

If third party users have a license, they should follow the below instructions:

**1.** After got the license and downloaded the database which is in a xml file, in order to run the program, the xml file has to be put in the same directory with **TkAstroDb.py** script file.

**2.** Before running the program make sure that **TkAstroDb** directory tree contains at least the following:

    TkAstroDb+
             |_Eph+
             |_adb_export_181128_2309.xml
             |_TkAstroDb.py

## Availability

Windows, Linux and MacOSX

## Dependencies

In order to run **TkAstroDb**, at least [Python](https://www.python.org/)'s 3.6 version must be installed on your computer. Note that in order to use [Python](https://www.python.org/) on the command prompt, [Python](https://www.python.org/) should be added to the PATH. There is no need to install manually the libraries that are used by the program. When the program first runs, the necessary libraries will be downloaded and installed automatically.

## Usage

**1.** Run the program by writing the below to **cmd** for Windows or to **bash** for Unix. (Users should wait a bit the program to be opened. Because the program will try to find all records and categories from the xml file.)

**For Unix**

    python3 TkAstroDb.py

**For Windows**

    python TkAstroDb.py

**Note:** When the program first run in Windows, users will get a [PermissionError](https://docs.python.org/3/library/exceptions.html#PermissionError)  during the installation of [Pyswisseph](https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyswisseph) library unless they run **cmd** as Administrator.

**2.** Users should see a window after 10-15 minutes which is similar to below. 

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

**7.** Users can focus on a record then by right clicking to that record they can see some options that can be done with the focused record. One of the options is deleting the record from displayed results. The deleted record will not be used in the computation later. The other option opens the [Astro-Databank](https://www.astro.com/astro-databank/Main_Page) webpage of the selected record.

![img7](https://user-images.githubusercontent.com/29302909/55159437-10b61200-5172-11e9-8bd9-bce0a526623b.png)

**8.** If users click **Export** menu button, they can see **Adb Links**, **Latitude Frequency** and **Year Frequency** options. By clicking **Adb Links** option, they can export the links of displayed records to **links.txt** file. This file will be created at **TkAstroDb** directory. By clicking the **Latitude Frequency** option, the latitude intervals of the selected categories and the mean latitude value will be written inside **latitude-frequency.txt** file. By clicking the **Year Frequency** option, a windows is opened as below. Users can specify the maximum, minimum and step values. After clicked the **Apply** button, the frequency of the years of displayed records will be put in a **year-frequency.txt** file.

![img8](https://user-images.githubusercontent.com/29302909/51045381-80394e00-15d4-11e9-8eed-881fa66f0afb.png)

**9.** If users click **Options** menu button, they can see **House System** option. And by clicking that menu button, they can define the house system they want to use. If they don't click this menu button, the house system will be defined according to the default setting. The default house system were defined as **Placidus**.

![img9](https://user-images.githubusercontent.com/29302909/51045506-d1e1d880-15d4-11e9-88f6-b63011273862.png)

**10.** If users click **Options** menu button, they can see **Orb Factor** option. And by clicking that menu button, they can define the orb factors for each astrological aspect. If users don't click this menu button, the orb factors will be defined according to their default settings. The default orb factors were defined as follow:

![img10](https://user-images.githubusercontent.com/29302909/50407124-2cf88a80-07e2-11e9-92f4-d51a4f7a6697.png)

**11.** If users click **Records** menu button, they can see an option which name is as **Create New Record**. And if users select that option, they will see a new window which is as follows:

![img11](https://user-images.githubusercontent.com/29302909/55160888-5a542c00-5175-11e9-8259-4054157d93dc.png)

**12.** Users should fill all the entry fields that can be seen on the image. They can select any existing categories or they can define new categories using entry field if they want. Finally users should press the **Add** button, then the selected or defined categories will be added to the category listbox. Note that users should use decimal latitude and longitude coordinates for a place.

The encoding of Windows is [cp1252](https://en.wikipedia.org/wiki/Windows-1252) while the encoding of Unix is [utf-8](https://en.wikipedia.org/wiki/UTF-8). That's why in Windows non-ASCII characters give an [UnicodeDecodeError](https://wiki.python.org/moin/UnicodeDecodeError). When users try to add a new record to the new database users will receive [UnicodeDecodeError](https://wiki.python.org/moin/UnicodeDecodeError) if the place that is found via latitude and longitude values, contains non-ASCII characters. This error occurs because of the codes of the [countryinfo](https://pypi.org/project/countryinfo/) library. That's why users should manually change some parts of [countryinfo](https://pypi.org/project/countryinfo/) library to avoid from this problem.

What users should do is simple:

1. Go to Python's site-packages library folder. For example, if Python3.6 or Python3.7 is installed on Program Files directory, users should go to below path:

       C:\Program Files\Python36\Lib\site-packages\countryinfo

2. Open the **countryinfo.py** script file.

3. Go to the 30'th line, the below codes on this line should be seen:

                country_info = json.load(open(file_path))

4. Replace the above code with below code:

                country_info = json.load(open(file_path, encoding="utf-8"))

5. Save and exit the script file. Now users no longer get an [UnicodeDecodeError](https://wiki.python.org/moin/UnicodeDecodeError) because of non-ASCII characters.

![img12](https://user-images.githubusercontent.com/29302909/55161048-b7e87880-5175-11e9-8915-771bf09e8bd8.png)

**13.** After added a new record to an alternative database which name is **TkAstroDb.db**, the new category that was defined by the user will be added to the category list. Then if users click the **Select** button which is near the **Categories** label, the newly defined category can be seen.

![img13](https://user-images.githubusercontent.com/29302909/55161277-2fb6a300-5176-11e9-8d0c-db32a4ac623e.png)

**14.** After clicked the **Apply** button and selected the Rodden Rating, the newly added record can be displayed.

![img14](https://user-images.githubusercontent.com/29302909/55178782-1eca5980-5197-11e9-912b-aed6b584531f.png)

**15.** If users click **Calculations** menu button, they can see options which are as follows: **Find Observed Values**, **Find Expected Values**, **Find Chi-Square Values** and **Find Effect Size Values**. If users want to find the astrological pattern distributions of any category, they should click **Find Observed Values** button. After clicked that menu button, a progress bar should be created as follows:

![img15](https://user-images.githubusercontent.com/29302909/55159732-cc774180-5172-11e9-88b8-14ad17ccd6c2.png)

**16.** After the computation finished, a log file (**output.log**) and an excel spreadsheet file (**observed_values.xlsx**) can be found inside nested directories like **Vocation/Occult_Fields/Astrologer/Rodden_Rating_AA/Orb_Factor_6_2_2_4_2_6_6_2_2_2_6/House_System_Placidus**. The directory names can be different according to the settings selected by the users. The spreadsheet file contains the astrological pattern distributions of displayed records.

![img16](https://user-images.githubusercontent.com/29302909/51046181-79133f80-15d6-11e9-8457-cc72a010e63f.png)

**17.** In order to calculate the expected values, the users must have two tables which include the astrological pattern distributions of two different categories. The expected values are calculated by comparing this two different categories. One category will be used as a *control group*, the other category will be used as a *research group*. While the table which is wanted to use as a *control group* should be renamed as **control_group.xlsx**, there is no need to change the name of *research group*, so it's name should be **observed_values.xlsx**. Note that users should copy the related tables to the **TkAstroDb** folder, then users can click **Calculations** menu button and they should select **Find Expected Values** option. There are two different methods to calculate the expected results.

**17.1. Flavia's Method:** This method is recommended to be used when the population number of the control group is small. For example this method can be used when the control group is another Adb category.
    
```python3
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
       
from random import randrange
       
       
def formula(x: list, y: list):
    return [sum(x) * (x[i] + y[i]) / (sum(x) + sum(y)) for i in range(12)],\
        [sum(y) * (x[i] + y[i]) / (sum(x) + sum(y)) for i in range(12)]
        
           
x1 = [randrange(0, 100, 1) for i in range(12)]
y1 = [randrange(0, 100, 1) for i in range(12)]

exp_x, exp_y = formula(x=x1, y=y1) 

print(f"x1: {sum(x1)}, Expected x1: {sum(exp_x)}\n\
y1: {sum(y1)}, Expected y1: {sum(exp_y)}")
```
       
**17.2. Sjoerd's Method:** This method is recommended to be used when the population number of the control group is larger.
       
```python3
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
       
from random import randrange

       
def formula(x: list, y: list):
    return [i * sum(x) / sum(y) for i in y], [i * sum(y) / sum(x) for i in x] 
        
         
x1 = [randrange(0, 100, 1) for i in range(12)]  
y1 = [randrange(0, 100, 1) for i in range(12)]

exp_x, exp_y = formula(x=x1, y=y1) 

print(f"x1: {sum(x1)}, Expected x1: {sum(exp_x)}\n\
y1: {sum(y1)}, Expected y1: {sum(exp_y)}")
```

**17.3.** After the calculation finished, **control_group.xlsx** and **observed_values.xlsx** files will be deleted and a new excel spreadsheet file will be created as **expected_values.xlsx**. And it is recommended to move **expected_values.xlsx** file to the related category folder:

![img17](https://user-images.githubusercontent.com/29302909/51046081-3b161b80-15d6-11e9-81d9-936c39d982e0.png)

**18.** In order to calculate the chi-square values, users should copy **expected_values.xlsx** and **observed_values.xlsx** files to **TkAstroDb** folder. Then they should select **Calculations** menu button. After that, they should click **Find Chi-Square Values** option. After the compuation finished, **expected_values.xlsx** and **observed_values.xlsx** files will be deleted and a new excel spreadsheed file will be created as **chi-square.xlsx** in **TkAstroDb** folder. It is recommended that the users cut the file from **TkAstroDb** directory then paste it to the related directory. And it is recommended to move **chi-square.xlsx** file to the related folder:

![img18](https://user-images.githubusercontent.com/29302909/51046111-48330a80-15d6-11e9-8e05-d3e21d619b61.png)

**19.** In order to calculate the effect size values, users should copy **expected_values.xlsx** and **observed_values.xlsx** files to the **TkAstroDb** directory. Then they should select **Calculations** menu button. After that, they should select **Find Effect Size Values** option. After the compuation finished, **expected_values.xlsx** and **observed_values.xlsx** will be deleted and a new excel spreadsheed file will be created as **effect-size.xlsx** in **TkAstroDb** folder. It is recommended that the users cut this file from **TkAstroDb** directory then paste it to the related directory. And it is recommended to move **effect-size.xlsx** file to the related folder:

![img19](https://user-images.githubusercontent.com/29302909/51046129-54b76300-15d6-11e9-9c18-ac49810666cb.png)

## Spreadsheets

**observed_values.xlsx**

![img20](https://user-images.githubusercontent.com/29302909/55176661-9944aa80-5192-11e9-8891-ca4498541e8b.png)

**expected_values.xlsx**

![img21](https://user-images.githubusercontent.com/29302909/55176770-d27d1a80-5192-11e9-95e9-9aee8916ccfb.png)

**chi-square.xlsx**

![img22](https://user-images.githubusercontent.com/29302909/55176879-0c4e2100-5193-11e9-923f-dfe0c3b4a934.png)

**effect-size.xlsx**

![img23](https://user-images.githubusercontent.com/29302909/55176960-369fde80-5193-11e9-979a-f7429821261d.png)

## Notes

**1.** The tables may not be opened by Microsoft Excel. Therefore it is recommended to use Libre Office. 

**2.** If users want to put their files in a cloud system like Dropbox, it is recommended that the format of the excel files should be changed from *xlsx* to *ods* format.

## Licenses

TkAstroDb is released under the terms of the GNU GENERAL PUBLIC LICENSE. Please refer to the LICENSE file.
