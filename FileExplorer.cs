using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Interfaces;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using static System.Collections.Specialized.BitVector32;

namespace GMRMS
{
    [TestClass]
    public class FileExplorer
    {
        System.Diagnostics.Process winAppDriverProcess;
        AppiumOptions options;
        WindowsDriver<WindowsElement> Fexplorer;


        [TestInitialize]
        public void Initialize()
        {
            // Initiate WinAppDriver
            winAppDriverProcess = System.Diagnostics.Process.Start(@"C:\\Program Files (x86)\\Windows Application Driver\\WinAppDriver.exe");

            // Launch Caliculator 
            options = new AppiumOptions();
            //options.AddAdditionalCapability("app", @"C:\gmTestBed\FileExplorer\src\FileExplorer.exe"); // VB6 Applicaiton
            options.AddAdditionalCapability("app", @"C:\gmTestBed\FileExplorer\proj_csh\deploy\FileExplorer\bin\FileExplorer.exe");

            Fexplorer = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), options);
        }


        [TestMethod]
        public void Exitthetool()
        {
            // Select file and exit fileexplorer
            Fexplorer.FindElementByName("File").Click();
            Thread.Sleep(2000);
            Fexplorer.FindElementByName("Exit").Click();

            Thread.Sleep(2000);

            //Fexplorer.FindElementByName("Exit").Click();

            //Fexplorer.FindElementByClassName("ListView20WndClass").
            //Fexplorer.FindElementByXPath("//*[@Name='File Explorer']/Pane[@ClassName='TreeView20WndClass']").Click();
            //Fexplorer.FindElementByClassName("ListView20WndClass").Click() ;
                        //Fexplorer.FindElementByName("appverifUI.dll").Click();
            //Fexplorer.FindElementByXPath("//*[@Name='File Explorer']/Pane[@ClassName='ListView20WndClass']").Click();
              
           Fexplorer.Close();

        }

        [TestMethod]
        public void FocusCD()
        {
            //Fexplorer.FindElementByAccessibilityId("Toolbar1").Click();
            
            /*
            Thread.Sleep(1000);
            Fexplorer.FindElementByName("up").Click();
            var txt = Fexplorer.FindElementByName("up").Text;
            Console.WriteLine(txt);
            */


            // down
            Fexplorer.FindElementByName("C:\\").Click();
            Fexplorer.FindElementByName("down").Click();
            Thread.Sleep(2000);

                 //var dtext = Fexplorer.FindElementByAccessibilityId("AddressBandRoot").Text;
                 //Console.WriteLine(dtext);

            var focusedElement = Fexplorer.SwitchTo().ActiveElement();
            var focusedDriveText = focusedElement.GetAttribute("Name");
            Console.WriteLine("Focused Drive:D " + focusedDriveText); // Print the text of the focused drive (D drive)
            Thread.Sleep(2000);

            if (focusedDriveText.Contains("D"))
            {
                Console.WriteLine(" Results Match - Focused D Drive ");

            }
            else
            {
                Console.WriteLine("Results unMatch - Not Focused D Drive");

            }

            // Up 
            Fexplorer.FindElementByName("up").Click();
            var focusedElement1 = Fexplorer.SwitchTo().ActiveElement();
            var focusedDriveText1 = focusedElement1.GetAttribute("Name");
            Console.WriteLine("Focused Drive:C  " + focusedDriveText1); // Print the text of the focused drive (C drive)

            if (focusedDriveText1.Contains("C"))
            {
                Console.WriteLine(" Results Match - Focused C Drive ");

            }
            else
            {
                Console.WriteLine("Results unMatch - Not Focused D Drive");

            }
            Fexplorer.Close();
        }

        [TestMethod]
        public void FileContent()
        {
            Fexplorer.FindElementByName("C:\\").Click();
            Thread.Sleep(1000);

            // Read the file content and print
            Fexplorer.FindElementByName("appverifUI.dll").Click();
            Thread.Sleep(1000);
            var adlltxt = Fexplorer.FindElementByName("RichEdit Control").Text;
            Console.WriteLine(adlltxt);

            //var lst = Fexplorer.FindElementsByAccessibilityId("lvListView");

            // Read the files in the right side of the Panal and print 
            var filesPanel = Fexplorer.FindElementByAccessibilityId("lvListView"); // Locate the panel
            var fileElements = filesPanel.FindElementsByTagName("ListItem"); // Find all elements within the panel
            List<string> fileNames = new List<string>(); // Create a list to store file names

            // Extract file names and add them to the list
            foreach (var fileElement in fileElements)
            {
                string fileName = fileElement.GetAttribute("Name");
                fileNames.Add(fileName);
            }

            // Print the list of file names
            foreach (var fileName in fileNames)
            {
                Console.WriteLine(fileName);
            }
            
            // Print no of files in the folder
            Console.WriteLine("No of files: {0} ", fileNames.Count);

            Fexplorer.Close();
        }


        [TestMethod]
        public void FileNamesSort()
        {
            
            // Clicking on the "C:\" drive
            Fexplorer.FindElementByName("C:\\").Click();

            // Read and extract the filename before sorting
            var filesPanel = Fexplorer.FindElementByAccessibilityId("lvListView"); // Locate the files in the panel
            var fileElementsBeforeSorting = filesPanel.FindElementsByTagName("ListItem"); // Find all fileNames 

            List<string> fileNamesBeforeSorting = new List<string>(); // Create a list to store file names 

            foreach (var fileElement in fileElementsBeforeSorting) // Extract file names and add them to the list
            {
                string fileName = fileElement.GetAttribute("Name");
                fileNamesBeforeSorting.Add(fileName);
            }
                        
            Console.WriteLine("File Names Before Sorting:");// Print the list of file names 
            foreach (var fileName in fileNamesBeforeSorting)
            {
                Console.WriteLine(fileName);
            }

            // Sorting 
            Fexplorer.FindElementByName("Name").Click(); // Click on the "Name" header to sort

            Thread.Sleep(4000);

            //Fexplorer.FindElementByName("Name").Click();
            //Thread.Sleep(4000);

            
            WindowsElement nameColumn = Fexplorer.FindElementByName("Name"); // Locate the "Name" column header

            // Double-click the "Name" column header
            Actions action = new Actions(Fexplorer);
            action.DoubleClick(nameColumn).Perform();
            Thread.Sleep(4000);

            //Fexplorer.FindElementByName("C:\\").Click();
            Fexplorer.FindElementByName("Name").Click();
            Thread.Sleep(2000);


            List<string> fileNamesAfterSorting = new List<string>(); // Create a list to store file names 

            Fexplorer.FindElementByClassName("WindowsForms10.SysListView32.app.0.141b42a_r8_ad1").Click();
            Thread.Sleep(4000);

            var filesPanel1 = Fexplorer.FindElementByAccessibilityId("lvListView");
            
            var fileElementsAfterSorting = filesPanel1.FindElementsByTagName("ListItem"); // Find all elements again after sorting


            foreach (var fileElement in fileElementsAfterSorting)  // Extract file names after sorting and add them to the list
            {
                string fileName = fileElement.GetAttribute("Name");
                fileNamesAfterSorting.Add(fileName);
            }

            
            Console.WriteLine("\nFile Names After Sorting:"); // Print the sorted list of file names
            foreach (var fileName in fileNamesAfterSorting)
            {
                Console.WriteLine(fileName);
            }



            // Compare the file names before and after sorting
            bool filesMatch = true;
            for (int i = 0; i < fileNamesBeforeSorting.Count; i++)
            {
                if (fileNamesBeforeSorting[i] != fileNamesAfterSorting[i])
                {
                    filesMatch = false;
                    break;
                }
            }

            if (filesMatch)
            {
                Console.WriteLine("\nFiles are not sorted correctly!");
            }
            else
            {
                Console.WriteLine("\nFiles are sorted correctly!");
            }

            // Close the session
            Fexplorer.Quit();
        }

        [TestMethod]
        public void CSubFolders()
        {
            // Select a sunfolder in Recycle Bin
            Fexplorer.FindElementByName("C:\\").Click();
            Fexplorer.FindElementByName("$Recycle.Bin").Click();
            Thread.Sleep(2000);
            Fexplorer.FindElementByName("S-1-5-18").Click(); // Select subfolder in RecycleBin folder
            Fexplorer.FindElementByName("Continue").Click();


            // Moving to folder down option
            Fexplorer.FindElementByName("C:\\").Click();
            Fexplorer.FindElementByName("$Recycle.Bin").Click();
            Fexplorer.FindElementByName("down").Click();
            Thread.Sleep(100);
            Fexplorer.FindElementByName("down").Click();

            var focus = Fexplorer.SwitchTo().ActiveElement();
            var focusText = focus.GetAttribute("Name");
            Console.WriteLine("Focused Folder: " + focusText); // Print the text of the focused folder
            Thread.Sleep(2000);


            // Moving to folder up option
           
            Fexplorer.FindElementByName("up").Click();

            var focus1 = Fexplorer.SwitchTo().ActiveElement();
            var focusText1 = focus1.GetAttribute("Name");
            Console.WriteLine("Focused Folder: " + focusText1); // Print the text of the focused folder
            Thread.Sleep(2000);

            Fexplorer.Close();
        }

        [TestMethod]
        public void OpenAndClose()
        {
            
            // Locate the tree view element that contains the folder list
            var treeView = Fexplorer.FindElementByAccessibilityId("tvTreeView");

            // Find all the child elements (folders) within the tree view
            var folderElements = treeView.FindElementsByClassName("WindowsForms10.SysTreeView32.app.0.141b42a_r8_ad1");

            List<string> folderNames = new List<string>();

            // Extract the names of the folders and add them to the list
            foreach (var folderElement in folderElements)
            {
                string folderName = folderElement.GetAttribute("Name");
                folderNames.Add(folderName);
            }

            // Print the list of folder names
            Console.WriteLine("Folder Names:");
            foreach (var folderName in folderNames)
            {
                Console.WriteLine(folderName);
            }


            // Verfify the Open and close options with Folder panel data

            Fexplorer.FindElementByName("C:\\").Click();
            Fexplorer.FindElementByName("close").Click();
            Thread.Sleep(2000);
            Fexplorer.FindElementByName("open").Click();

            //Take Screen shot after select Open
            var sshot = Fexplorer.GetScreenshot();
            string sspath = "C:\\Users\\ala\\Documents\\Open.png";
            sshot.SaveAsFile(sspath, ScreenshotImageFormat.Png);
            Console.WriteLine($"Screenshot saved to: {sspath}");
            Thread.Sleep(2000);
            
            //Click on Close
            Fexplorer.FindElementByName("close").Click();
            Thread.Sleep(2000);

            //Fexplorer.FindElementByName("open").Click();  

            //Take Screen shot after close
            var sshot1 = Fexplorer.GetScreenshot();
            string sspath1 = "C:\\Users\\ala\\Documents\\Close.png";
            sshot1.SaveAsFile(sspath1, ScreenshotImageFormat.Png);
            Console.WriteLine($"Screenshot saved to: {sspath1}");
           
            /*
            // Compare the screen shot before and after 
            if (sshot == sshot1)
            {
                Console.WriteLine("Screenshots are identical");
            }
            else
            {
                Console.WriteLine("Screenshots are different");
            }

            */

            Bitmap sshot3 = new Bitmap("C:\\Users\\ala\\Documents\\Open.png");

            Bitmap sshot4 = new Bitmap("C:\\Users\\ala\\Documents\\Close.png");



            if (sshot3.Width == sshot4.Width && sshot3.Height == sshot4.Height)
            {
                bool imagesAreEqual = true;

                // Compare pixel by pixel
                for (int x = 0; x < sshot3.Width && imagesAreEqual; x++)
                {
                    for (int y = 0; y < sshot3.Height && imagesAreEqual; y++)
                    {
                        if (sshot3.GetPixel(x, y) != sshot4.GetPixel(x, y))
                        {
                            imagesAreEqual = false;
                        }
                    }
                }

                if (imagesAreEqual)
                {
                    Console.WriteLine("Screenshots are identical.");
                }
                else
                {
                    Console.WriteLine("Screenshots are different.");
                }


                Fexplorer.Close();

            }

        }
    }
}
