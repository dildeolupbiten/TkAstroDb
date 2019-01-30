# TkAstroDb

TkAstroDb is a program that uses Astrodatabank (Adb) to conduct statistical studies in astrology. Because of the license conditions, Astrodatabank can not be shared, copied with third party users. Therefore those who are interested in using that program should contact with the webmaster of http://www.astro.com to get a license.

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

In order to run **TkAstroDb**, at least [Python](https://www.python.org/)'s 3.6 version must be installed on your computer. Note that in order to use Python on the command prompt, Python should be added to the PATH. There is no need to install manually the libraries that are used by the program. When the program first runs, the necessary libraries will be downloaded and installed automatically.

## Usage

**1.** Run the program by writing the below to **cmd** for Windows or to **bash** for Unix. (Users should wait a bit the program to be opened. Because the program will try to find all records and categories from the xml file.)

**For Unix**

    python3 TkAstroDb.py

**For Windows**

    python TkAstroDb.py

**Note:** When the program first run in Windows, users will get a **Permission Error**  during the installation of **Pyswisseph** library unless they run the **cmd** as Administrator.

**2.** Users should see a window after 10-15 minutes which is similar to below. 

![img1](https://user-images.githubusercontent.com/29302909/51044209-985b9e00-15d1-11e9-8e7e-c132928287a2.png)

**3.** If users want to add single records to the displayed records according to the selection, they should type the name of the record. For example suppose a user wants to add **Albert Einstein** to the displayed records, the user should write **Einstein, Albert** to the **Search A Record By Name** section. While typing the name, if the record is found, the user will see an **Add** button which is used for adding records to the treeview.

![img2](https://user-images.githubusercontent.com/29302909/51044229-a9a4aa80-15d1-11e9-98e1-7dbbd4fd7e53.png)

![img3](https://user-images.githubusercontent.com/29302909/51044247-b4f7d600-15d1-11e9-93cb-1634bd008519.png)

**4.** If users right click on the search entry field, the following options will be opened in the right click menu. 

![img4](https://user-images.githubusercontent.com/29302909/51044263-c214c500-15d1-11e9-850e-ecfa1743f6f1.png)

**5.** Click **Select** button which is near to **Categories** label. After that a small window should be opened as below. Select one category or more categories to study, then press **Apply** button.

![img5](https://user-images.githubusercontent.com/29302909/50767418-d1bca280-128d-11e9-9c30-a7ce39884f35.png)

**6.** Click **Select** button which is near to **Rodden Rating** label. A small window should be opened as below. Select one rating or more ratings then press **Apply** button.

![img6](https://user-images.githubusercontent.com/29302909/50359618-72148680-056e-11e9-94e0-17938c41d268.png)

**7.** Before clicking **Display Records** button, users can click the check buttons in order to filter the records. Because some human records can be in event categories or some event records can be in human categories. But if users want to display all records, they should not click the check buttons. Then, users can click **Display Records** button. After clicked to that button, users should wait a bit. Finally the records will be displayed at treeview as follows:

![img7](https://user-images.githubusercontent.com/29302909/51044370-043e0680-15d2-11e9-8cf5-f63edde5660b.png)

**8.** Users can focus on a record then by right clicking to that record they can see some options that can be done with the focused record. One of the options is deleting the record from displayed results. The deleted record will not be used in the computation later. The other option opens the Adb webpage of that record.

![img8](https://user-images.githubusercontent.com/29302909/51044409-1455e600-15d2-11e9-8895-5d007c9b64c5.png)

**9.** If users click **Export** menu button, they can see **Adb Links** and **Year Frequency** options. By clicking **Adb Links** option, they can export the links of displayed records to **links.txt** file. This file will be created at **TkAstroDb** directory. By clicking the **Year Frequency** option, a windows is opened as below. Users can specify the maximum, minimum and step values. After clicked the **Apply** button, the frequency of the years of displayed records will be put in a **year-frequency.txt** file.

![img9](https://user-images.githubusercontent.com/29302909/51045381-80394e00-15d4-11e9-8eed-881fa66f0afb.png)

**10.** If users click **Options** menu button, they can see **House System** option. And by clicking that menu button, they can define the house system they want to use. If they don't click this menu button, the house system will be defined according to the default setting. The default house system were defined as **Placidus**.

![img10](https://user-images.githubusercontent.com/29302909/51045506-d1e1d880-15d4-11e9-88f6-b63011273862.png)

**11.** If users click **Options** menu button, they can see **Orb Factor** option. And by clicking that menu button, they can define the orb factors for each astrological aspect. If users don't click this menu button, the orb factors will be defined according to their default settings. The default orb factors were defined as +-6 for Conjunction, +-2 for Semi-Sextile, +-2 Semi-Square, +-4 for Sextile, +-2 for Quintile, +-6 for Square, +-6 for Trine, +-2 for Sesquiquadrate, +-2 for BiQuintile, +-2 for Quincunx and +-6 for Opposite aspects.

![img11](https://user-images.githubusercontent.com/29302909/50407124-2cf88a80-07e2-11e9-92f4-d51a4f7a6697.png)

**12.** If users click **Calculations** menu button, they can see options which are as follows: **Find Observed Values**, **Find Expected Values**, **Find Chi-Square Values** and **Find Effect Size Values**.. If users want to find the astrological pattern distributions of any category, they should click **Find Observed Values** button. After clicked that menu button, a progress bar should be created as follows:

![img12](https://user-images.githubusercontent.com/29302909/51044436-23d52f00-15d2-11e9-8dcd-f51f244c09b1.png)

**13.** After the computation finished, a log file (**output.log**) and an excel spreadsheet file (**observed_values.xlsx**) can be found inside nested directories like **Vocation/Occult_Fields/Astrologer/Rodden_Rating_AA/Orb_Factor_6_2_2_4_2_6_6_2_2_2_6/House_System_Placidus**. The directory names can be different according to the settings selected by the users. The spreadsheet file contains the astrological pattern distributions of displayed records.

![img13](https://user-images.githubusercontent.com/29302909/51046181-79133f80-15d6-11e9-8457-cc72a010e63f.png)

**14.** In order to calculate the expected values, the users must have two tables which include the astrological pattern distributions of two different categories. The expected values are calculated by comparing this two different categories. One category will be used as a *control group*, the other category will be used as a *research group*. While the table which is wanted to use as a *control group* should be renamed as **control_group.xlsx**, there is no need to change the name of *research group*, so it's name should be **observed_values.xlsx**. Note that users should copy the related tables to the **TkAstroDb** folder, then users can click **Calculations** menu button and they should select **Find Expected Values** option. There are two different methods to calculate the expected results.

**14.1. Flavia's Method:** This method is recommended to be used when the population number of the control group is small. For example this method can be used when the control group is another Adb category.
    
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
       
**14.2. Sjoerd's Method:** This method is recommended to be used when the population number of the control group is larger.
       
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
    

**14.3.** After the calculation finished, **control_group.xlsx** and **observed_values.xlsx** files will be deleted and a new excel spreadsheet file will be created as **expected_values.xlsx**. And it is recommended to move **expected_values.xlsx** file to the related category folder:

![img14](https://user-images.githubusercontent.com/29302909/51046081-3b161b80-15d6-11e9-81d9-936c39d982e0.png)

**15.** In order to calculate the chi-square values, users should copy **expected_values.xlsx** and **observed_values.xlsx** files to **TkAstroDb** folder. Then they should select **Calculations** menu button. After that, they should click **Find Chi-Square Values** option. After the compuation finished, **expected_values.xlsx** and **observed_values.xlsx** files will be deleted and a new excel spreadsheed file will be created as **chi-square.xlsx** in **TkAstroDb** folder. It is recommended that the users cut the file from **TkAstroDb** directory then paste it to the related directory. And it is recommended to move **chi-square.xlsx** file to the related folder:

![img15](https://user-images.githubusercontent.com/29302909/51046111-48330a80-15d6-11e9-8e05-d3e21d619b61.png)

**16.** In order to calculate the effect size values, users should copy **expected_values.xlsx** and **observed_values.xlsx** files to the **TkAstroDb** directory. Then they should select **Calculations** menu button. After that, they should select **Find Effect Size Values** option. After the compuation finished, **expected_values.xlsx** and **observed_values.xlsx** will be deleted and a new excel spreadsheed file will be created as **effect-size.xlsx** in **TkAstroDb** folder. It is recommended that the users cut this file from **TkAstroDb** directory then paste it to the related directory. And it is recommended to move **effect-size.xlsx** file to the related folder:

![img16](https://user-images.githubusercontent.com/29302909/51046129-54b76300-15d6-11e9-9c18-ac49810666cb.png)

## Spreadsheets

**observed_values.xlsx**

![img17](https://user-images.githubusercontent.com/29302909/51140511-cc47f500-1857-11e9-9867-ae413c28c379.png)

**expected_values.xlsx**

![img18](https://user-images.githubusercontent.com/29302909/51140623-103afa00-1858-11e9-9a8b-c7862cfe6654.png)

**chi-square.xlsx**

![img19](https://user-images.githubusercontent.com/29302909/51140580-f0a3d180-1857-11e9-8c4f-a81b37974b4a.png)

**effect-size.xlsx**

![img20](https://user-images.githubusercontent.com/29302909/51140744-6445de80-1858-11e9-883a-3c68e7dbe97d.png)

## Notes

**1.** The tables may not be opened by Microsoft Excel. Therefore it is recommended to use Libre Office. 

**2.** If users want to put their files in a cloud system like Dropbox, it is recommended that the format of the excel files should be changed from *xlsx* to *ods* format.

## Licenses

TkAstroDb is released under the terms of the GNU GENERAL PUBLIC LICENSE. Please refer to the LICENSE file.
