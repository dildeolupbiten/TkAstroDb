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

In order to run **TkAstroDb**, at least Python's 3.6 version must be installed on your computer. Note that in order to use Python on the command prompt, Python should be added to the PATH. Users don't have to install manually the libraries that are used by the program, when the program first runs, the necessary libraries will be downloaded and installed automatically.

## Usage

**1.** Run the program as follows. (Users should wait a bit the program to be opened. Because the program will try to find all records and categories from the xml file.)

    python3 TkAstroDb.py

or if the above command gives an error on Windows, please run the program as follows:

    python TkAstroDb.py

**Note:** If Windows users get a **Permission Error**  while the installation of **Pyswisseph** library, they should open the **cmd** as Administrator, then they should retype the above commands. Note that the upper directories of **TkAstroDb** directory should not contain spaces.

**2.** After that, a window should be opened which is similar to the below. 

![img1](https://user-images.githubusercontent.com/29302909/50402991-0f103300-07ac-11e9-98e1-fab84856cd47.png)

**3.** If users want to add single records to the displayed records according to the selection, they should type the name of the record. For example suppose a user wants to add **Albert Einstein** to the displayed records, the user should write **Einstein, Albert** to the **Search A Record By name** section. While typing the name, if the record is found, the user can add that record to the treeview by clicking the **Add** button.

![img2](https://user-images.githubusercontent.com/29302909/50403025-729a6080-07ac-11e9-85a5-a83d45b76661.png)

![img3](https://user-images.githubusercontent.com/29302909/50403053-9d84b480-07ac-11e9-828b-03095fe4232d.png)

**4.** If users right click on the search entry field, the following options will be opened in the right click menu. 

![img4](https://user-images.githubusercontent.com/29302909/50403086-db81d880-07ac-11e9-91e7-53d65e7d99df.png)

**5.** Click **Select** button which is near to **Categories** label. After that a small window should be opened as below. Select one category or more categories to study, then press **Apply** button.

![img5](https://user-images.githubusercontent.com/29302909/50767418-d1bca280-128d-11e9-9c30-a7ce39884f35.png)

**6.** Click **Select** button which is near to **Rodden Rating** label. A small window should be opened as below. Select one rating or more ratings then press **Apply** button.

![img6](https://user-images.githubusercontent.com/29302909/50359618-72148680-056e-11e9-94e0-17938c41d268.png)

**7.** Before clicking **Display Records** button, users can click the check buttons in order to filter the records. Because some human records can be in event categories or some event records can be in human categories. But if users want to display all records, they should not click the check buttons. Then, users can click **Display Records** button. After clicked to that button, users should wait a bit. Finally the records will be displayed at treeview as follows:

![img7](https://user-images.githubusercontent.com/29302909/50403111-0ff59480-07ad-11e9-83dd-07d6d0d4e14a.png)

**8.** Users can focus on a record then by right clicking to that record they can see some options that can be done with the focused record. One of the options is deleting the record from displayed results. The deleted record will not be used in the computation later. The other option opens the Adb webpage of that record.

![img8](https://user-images.githubusercontent.com/29302909/50403125-37e4f800-07ad-11e9-9482-0e6e314b940e.png)

**9.** If users click **Export** menubutton, they can see **Adb Links** option. By clicking **Adb Links** option, they can export the links of displayed results to **links.txt** file. This file will be created at **TkAstroDb** directory.

**10.** If users click **Options** menu button, they can see **House System** option. And by clicking that menu button, they can define the house system they want to use. If they don't click this menu button, the house system will be defined according to the default setting. The default house system were defined as **Placidus**.

![img9](https://user-images.githubusercontent.com/29302909/50113149-29995800-0252-11e9-8563-b05f373aa0d4.png)

**11.** If users click **Options** menubutton, they can see **Orb Factor** option. And by clicking that menu button, they can define the orb factors for each astrological aspect. If users don't click this menu button, the orb factors will be defined according to their default settings. The default orb factors were defined as +-6 for Conjunction, +-2 for Semi-Sextile, +-2 Semi-Square, +-4 for Sextile, +-2 for Quintile, +-6 for Square, +-6 for Trine, +-2 for Sesquiquadrate, +-2 for BiQuintile, +-2 for Quincunx and +-6 for Opposite aspects.

![img10](https://user-images.githubusercontent.com/29302909/50407124-2cf88a80-07e2-11e9-92f4-d51a4f7a6697.png)

**12.** If users click **Calculations** menubutton, they can see options which are as follows: **Find Observed Values**, **Find Expected Values**, **Find Chi-Square Values** and **Find Effect Size Values**.. If users want to find the astrological pattern distributions of any category, they should click **Find Observed Values** button. After clicked that menu button, a progress bar should be created as follows:

![img11](https://user-images.githubusercontent.com/29302909/50379493-f8f35d00-065b-11e9-8c5c-361096cc1579.png)

**13.** After the computation finished, an excel spreadsheet will be created as **observed_values.xlsx** in nested directories like **Rodden_Rating_AA/Orb_Factor_6_2_2_4_2_6_6_2_2_2_6/House_System_Placidus**. The directory names can be different according to the settings that users selected. This file include the astrological pattern distributions of displayed records. Since third party users may want to find the pattern distributions of different categories, it is recommended to create a folder which name is **Categories**. And previously created excel file can be moved to that folder. Below can be seen a sample of the directory tree which is recommended. **Vocation** is the first category name. And there may be a lot of subcategories in that category. **Occult_Fields** is a subcategory of **Vocation** category. And **Astrologer** is the subcategory of **Occult_Fields** category. Thus users can know that the results came from the Rodden Rating AA. So there can be many sub directories in **Occult_Fields** category which relates to the different Rodden Ratings. Normally the default house system is **Placidus.** So **House_System_Placidus** directory is a sub directory of **Orb_Factor_6_2_2_4_2_6_6_2_2_2_6**. There can be many sub directories in **Orb_Factor_6_2_2_4_2_6_6_2_2_2_6** directory which relate to different house systems. **Orb_Factor_6_2_2_4_2_6_6_2_2_2_6** directory is in the **Rodden_Rating_AA** directory. The orb factors of the aspects can also be different and the orb factors can be defined by the user. So in **Rodden_Rating_AA** directory, there can be many sub directories which relates to different orb factors and different house systems. **links.txt** file includes the links of the records that are displayed on the treeview. And as can be seen, **observed_values.xlsx** file is in the **House_System_Placidus** directory.

![img12](https://user-images.githubusercontent.com/29302909/50403168-71b5fe80-07ad-11e9-90a3-ff1015787932.png)

**14.** In order to calculate the expected values, the users must have two tables which include the astrological pattern distributions of two different categories. The expected values are calculated by comparing this two different categories. One category will be used as a *control group*, the other category will be used as a *research group*. While the table which is wanted to use as a *control group* should be renamed as **control_group.xlsx**, there is no need to change the name of *research group*, so it's name should be **observed_values.xlsx**. Note that users should copy the related tables to the **TkAstroDb** folder. After that users can click **Calculations** menubutton and they should select **Find Expected Values** option. There are two different methods to calculate the expected results. Users can select any of them. After the calculation finished, **control_group.xlsx** and **observed_values.xlsx** files will be deleted and a new excel spreadsheet file will be created as **expected_values.xlsx**. And it is recommended to move **expected_values.xlsx** file to the related category folder:

![img13](https://user-images.githubusercontent.com/29302909/50427153-6ea43680-08b1-11e9-9b56-672613b507b8.png)

**15.** In order to calculate the chi-square values, users should copy **expected_values.xlsx** and **observed_values.xlsx** files to **TkAstroDb** folder. Then they should select **Calculations** menubutton. After that, they should click **Find Chi-Square Values** option. After the compuation finished, **expected_values.xlsx** and **observed_values.xlsx** files will be deleted and a new excel spreadsheed file will be created as **chi-square.xlsx** in **TkAstroDb** folder. It is recommended that the users cut the file from **TkAstroDb** directory then paste it to the related directory. And it is recommended to move **chi-square.xlsx** file to the related folder:

![img14](https://user-images.githubusercontent.com/29302909/50427148-3866b700-08b1-11e9-834c-4787196c7604.png)

**16.** In order to calculate the effect size values, users should copy **expected_values.xlsx** and **observed_values.xlsx** files to the **TkAstroDb** directory. Then they should select **Calculations** menubutton. After that, they should select **Find Effect Size Values** option. After the compuation finished, **expected_values.xlsx** and **observed_values.xlsx** will be deleted and a new excel spreadsheed file will be created as **effect-size.xlsx** in **TkAstroDb** folder. It is recommended that the users cut this file from **TkAstroDb** directory then paste it to the related directory. And it is recommended to move **effect-size.xlsx** file to the related folder:

![img15](https://user-images.githubusercontent.com/29302909/50427129-c68e6d80-08b0-11e9-962d-16e25a55570e.png)

## Notes

**1.** The tables may not be opened by Microsoft Excel. Therefore it is recommended to use Libre Office. 

**2.** If users want to put their files in a cloud system like Dropbox, it is recommended that the format of the excel files should be changed from *xlsx* to *ods* format.

## Licenses

TkAstroDb is released under the terms of the GNU GENERAL PUBLIC LICENSE. Please refer to the LICENSE file.

