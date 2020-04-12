using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace FindAndReplaceHelper
{
    class Program
    {
        static void Main(string[] args)
        { // TODO: use condition logic to elect a different starting paragraph for varied role types aka admin, support, dev, game dev next week.
            // maybe create loops on each readline assignment to prompt for confirmation
            // Change this to an exe file, or a WinFOrms GUI/Web App

            // implement try catch to avoid application crashes

            // TODO: Add unit tests
            // TODO: given option to select manual or automatic for the input and output paths
            // TODO: method to maybe increase length of what console.readline
            byte[] inputBuffer = new byte[1024];
            Stream inputStream = Console.OpenStandardInput(inputBuffer.Length);
            Console.SetIn(new StreamReader(inputStream, Console.InputEncoding, false, inputBuffer.Length));

            // Create new word application
            Application word = new Application();
            //word.Visible = true; // opens the doc
            // load all the MS Word data
            object miss = System.Reflection.Missing.Value;
            // object path = @"M:\//Troydon/Documents/Troydon/JobStuff/JobHunt/cover letters/Customs/SEEK_Free_cover_letter_template_2018_NZ.docx";
            Console.WriteLine("Please enter the file path you would like to use for the template.\nE.g.M:\\/Troydon/Documents/Troydon/JobStuff/JobHunt/cover letters/Customs/SEEK_Free_cover_letter_template_2018_NZ.docx");
            object path = Console.ReadLine() ?? string.Empty;
            object readOnly = true;
            Document docs = word.Documents.Open(ref path, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

            // TODO: if possible execute the shutdown word method to execute if compilation shuts unexpectedly or cancelled to avoid it being left opened.
            //if (Environment.Exit(exitCode))

            try
            {
                //
                // determine outputfile name
                object outputFI = new object(); 
                string output = "", inputError = "\nSorry you did not enter a valid input\n", hiringManagersName = null, relevantSkills = "", company = "", bonusQuestion = "", roleTitle = "";
                int count = 0;
                bool outputMatch = true;
                List<string> outputPath = new List<string>();

                do
                {
                    // Set user secrets for relevant skills 

                    // evaluates to true if none matches, false if there is a match ( anytime the loop is reset, we reassgin the correct output return type
                    outputMatch = CheckOutput(output);

                    // if we at count 0 we will ask the first question
                    if (count == 0)
                    {
                        Console.WriteLine("\nHit F1 to start, What output would you like: 'admin' 'itadmin' 'dev' 'support' 'sales'\n");

                        // Ensure we pressed F1 to continue, otherwise we start again, add this in front of every read line
                        ConsoleKeyInfo KP = Console.ReadKey();
                        if (!(KP.Key == ConsoleKey.F1))
                        {
                            Console.WriteLine("\nPlease hit F1 before typing");
                            count = 0;
                            outputMatch = true; // make it true so we can keep looping
                            continue;
                        }

                        output = Console.ReadLine();

                        // TODO: Next week -> Show preview of answers, instead of F1 show a preview of answer, hit enter again to move to next step if happy with answer, otherwise redo step by htting 'Esc' key
                        // this methods ensure that when we exit we close the background process as opposed to manually shutting down console
                        ExitApp(ref word, ref miss, ref docs, output);

                        // if we do not enter a valid output, the bool will be true and we exit the loop using continue
                        if (CheckOutput(output))
                        {
                            Console.WriteLine(inputError);
                            outputMatch = true; // make it true so we can keep looping
                            count = 0;

                            continue;
                        }

                        // if we  entered a valid input then we do something
                        //if (!CheckOutput(output))
                        if (output == "admin")
                        {
                            Console.WriteLine("Success");
                            relevantSkills = "I have three years’ work experience working in a call centre as a Customer Service Rep with Woolworths Mobile. Some of the duties of the role included managing SAP tickets, whereby I had to administer events such as network incidents, transport and logistics, customer complaints, and input relevant data such as customer information, device details, funds, network incidents and faults. I was also known by the company to perform well with these duties.";
                            count = 1;
                        }
                        else if (output == "itadmin")
                        {
                            relevantSkills = "I have relevant skills useful to the role, I have achieved this when I completed a Certificate IV in proramming, this includes RDBMS such as SQL Server, MySQL, SQL, T-SQL. I have three years’ work experience working in a call centre as a Customer Service Rep with Woolworths Mobile. Some of the duties of the role included managing SAP tickets, whereby I had to administer events such as network incidents, transport and logistics, customer complaints, and input relevant data such as customer information, device details, funds, network incidents and faults. I was also known by the company to perform well with these duties.";
                            count = 1;
                        }
                        else if (output == "dev")
                        {
                            relevantSkills = "I have relevant skills useful to the role, which I obtained when I completed a Certificate IV in programming, that include basic C# OOP, ASP.NET Core, HTML, CSS, JavaScript, jQuery, SQL DBMS, Unity3D, Windows 10 OS, Microsoft Office 365 including Word and Excel, basic understanding of cloud services such as Azure and AWS. I am now applying these technologies into building my knowledge and skill level in creating beautiful, user friendly software, web and game applications. I am also working on using the programming, scripting and mark-up languages to learn the programming concepts such as algorithms, data structures, writing user-centric functional specifications, writing scalable code, understanding conditional logic, database design, responsive design.";
                            count = 1;
                        }
                        else if (output == "support")
                        {
                            relevantSkills = "I am interested in this position; I believe I have the skills and enthusiasm needed to do well. I am looking to secure a role that involves working with technology. I enjoy working with technology and dealing with computers, and Windows OS. I have knowledge in Office 364 Suite and great troubleshooting skills, which I have gained when working with Woolworths Mobile as a Tech Support Representative for mobile devices including devices such as Android, iOS and OPPO.";
                            count = 1;
                        }
                        else if (output == "sales")
                        {
                            relevantSkills = "I am interested in this position; I believe I have the skills and enthusiasm needed to do well. I am looking to secure a role that involves working with technology. I enjoy working with technology and dealing with computers, and Windows OS. I have knowledge in Office 364 Suite and great sales skills, which I have gained when working with Woolworths Mobile as a Tech Support Representative for mobile devices including devices such as Android, and OPPO.";
                            count = 1;
                        }
                        
                    }

                    // once we have gathered details we again reassign the correct outputmatch value for false, each time the loop resets by assigning outputmatch to true to correct a previous step
                    outputMatch = CheckOutput(output);


                    // TODO: Maybe implement confirmation after each prompt later
                    // Gather all variables that we will need to insert into the document
                    string date = DateTime.Now.ToString("dddd, dd MMMM yyyy");

                    if (count == 1)
                    {
                        Console.WriteLine("\nHit F1 ti begin, then Enter the Hiring Manger's Name (enter 'n' if none)\n");

                        // Ensure to hit F1 or we move back to asking Hiring Managers Name
                        ConsoleKeyInfo KP = Console.ReadKey();
                        if (!(KP.Key == ConsoleKey.F1))
                        {
                            Console.WriteLine("\nPlease hit F1 before typing\n"); // error occurs when we do not press the correct key
                            count = 1;
                            outputMatch = true; // make it true so we can keep looping
                            continue;
                        }
                        hiringManagersName = Console.ReadLine();
                        ExitApp(ref word, ref miss, ref docs, hiringManagersName);
                        count = 2;

                        if (hiringManagersName == "n")
                        {
                            hiringManagersName = ""; // if there is no name we reassign the value with an empty string
                        }
                    }


                    Console.WriteLine("\nEnter F1 to start, Enter the company (hit 'n' if none)\n");

                    if (count == 2)
                    {
                        ConsoleKeyInfo KP = Console.ReadKey();
                        if (!(KP.Key == ConsoleKey.F1))
                        {
                            Console.WriteLine("\nPlease hit F1 before typing\n"); // error occurs when we do not press the correct key
                            count = 2;
                            outputMatch = true; // make it true so we can keep looping
                            continue;
                        }

                        company = Console.ReadLine();
                        ExitApp(ref word, ref miss, ref docs, company);
                        count = 3;
                        if (company == "n") company = ""; // if there is no name we reassign the value with an empty string

                    }

                    string addressingEmployer = ""; // initial value will be Dear Sir/madam
                    if (hiringManagersName == "") addressingEmployer = "Sir/Madam"; // if response was  no, we will assign with the name in place of Sir/Madam
                    else if (hiringManagersName != "") addressingEmployer = hiringManagersName;

                    if (count == 3)
                    {
                        Console.WriteLine("\nEnter F1 to start, then Enter the role title\n");

                        ConsoleKeyInfo KP = Console.ReadKey();
                        if (!(KP.Key == ConsoleKey.F1))
                        {
                            Console.WriteLine("\nPlease hit F1 before typing\n"); // error occurs when we do not press the correct key
                            count = 3;
                            outputMatch = true; // make it true so we can keep looping
                            continue;
                        }

                        roleTitle = Console.ReadLine();
                        ExitApp(ref word, ref miss, ref docs, roleTitle);
                        count = 4;
                    }

                    string advertiser = "Seek.com.au";
                    if (company != "n") advertiser = company; // if response was not no, we will assign with the name in place of <Company> 


                    if (count == 4)
                    {
                        Console.WriteLine("\nEnter F1 to start, Is there a bonus question?, type 'n' if none\n");

                        ConsoleKeyInfo KP = Console.ReadKey();
                        if (!(KP.Key == ConsoleKey.F1))
                        {
                            Console.WriteLine("Please hit F1 before typing\n"); // error occurs when we do not press the correct key
                            count = 4;
                            outputMatch = true; // make it true so we can keep looping
                            continue;
                        }

                        bonusQuestion = Console.ReadLine();
                        ExitApp(ref word, ref miss, ref docs, bonusQuestion);
                        count = 5;
                        if (bonusQuestion == "n") bonusQuestion = "";

                    }

                    //Console.WriteLine("Key-Skills. Here you want to highlight some of your core skills that talk to the key selection criteria for the role. E.g HTML");
                    //string keySkills = Console.ReadLine();

                    //Console.WriteLine("Talk about why you would like to work for the company and why you’d be a good fit. For example: Company name has been of interest" +
                    //    " to me since embarking on its mega store approach to retail. This is ideal for 21st century sales of flooring products. I was also impressed with" +
                    //    " the profile of your managing director Rod Smythe, which I read in the Retail journal late last year.");
                    //string whyWorkForCompany = Console.ReadLine();

                    // Replace text now, loop through this 8 times replacing all the needed text
                    for (int i = 0; i < 8; i++)
                    {
                        Find contentReplace = word.Selection.Find;
                        contentReplace.ClearFormatting();
                        Range rng = docs.Content;

                        // based on the iteration number we set what the text that is to be replaced will be
                        switch (i)
                        {
                            case 0:
                                contentReplace.Text = "<dd Month YYYY>";
                                break;
                            case 1:
                                contentReplace.Text = "<Hiring manager’s name>";
                                break;
                            case 2:
                                contentReplace.Text = "<Company>";
                                break;
                            case 3:
                                contentReplace.Text = "Sir/Madam";
                                break;
                            case 4:
                                contentReplace.Text = "<insert role title>";
                                break;
                            case 5:
                                contentReplace.Text = "Seek.com.au";
                                break;
                            case 6:
                                contentReplace.Text = "<bonus question>";
                                break;
                            case 7:
                                contentReplace.Text = "<insert relevant skills intro here>"; // might be skipping because it is to enter before other answers? if it does not work then put both out of loop
                                break;

                            //case 6:
                            //    contentReplace.Text = "<Key-Skills. Here you want to highlight some of your core skills that talk to the key selection criteria for the role. E.g HTML>";
                            //    break;
                            //case 7:
                            //    contentReplace.Text = "<Talk about why you would like to work for the company and why you’d be a good fit>";
                            //    break;
                            default:
                                break;
                        }

                        contentReplace.Replacement.ClearFormatting();
                        //contentReplace.Replacement.Text = date;

                        // depending on the text we are replacing we assign a different value for the replacement
                        switch (contentReplace.Text)
                        {
                            case "<dd Month YYYY>":
                                contentReplace.Replacement.Text = date;
                                break;
                            case "<Hiring manager’s name>":
                                contentReplace.Replacement.Text = hiringManagersName;
                                break;
                            case "<Company>":
                                contentReplace.Replacement.Text = company;
                                break;
                            case "Sir/Madam":
                                contentReplace.Replacement.Text = addressingEmployer;
                                break;
                            case "<insert role title>":
                                contentReplace.Replacement.Text = roleTitle; // TODO: make only this bold
                                break;
                            case "Seek.com.au":
                                contentReplace.Replacement.Text = advertiser;
                                break;
                            case "<bonus question>":
                                contentReplace.Replacement.Text = bonusQuestion;
                                break;
                            case "<insert relevant skills intro here>":
                                // if there is more than 250 characters in the parameter the replacement.text will not work so we implement copy and replace
                                if (relevantSkills.Length >= 250)
                                {
                                    contentReplace.Replacement.Text = "^c"; // copy to clipboard action is assigned to replacement
                                    contentReplace.Replacement.ClearFormatting();
                                    // now we search the whole document for this again
                                    rng.Find.Execute("<insert relevant skills intro here>", ref miss, ref miss, ref miss, ref miss, ref miss,
                                        ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss); // find this text
                                    // now replace the text in the document
                                    rng.Text = relevantSkills;
                                }
                                break;

                            //case "<Key-Skills. Here you want to highlight some of your core skills that talk to the key selection criteria for the role. E.g HTML>":
                            //    contentReplace.Replacement.Text = keySkills;
                            //    break;
                            //case "<Talk about why you would like to work for the company and why you’d be a good fit>":
                            //    contentReplace.Replacement.Text = whyWorkForCompany;
                            //    break;
                            default:
                                break;
                        }

                        // now execute the replace for all values
                        object replaceAll = WdReplace.wdReplaceAll;
                        contentReplace.Execute(ref miss, ref miss, ref miss, ref miss, ref miss,
                            ref miss, ref miss, ref miss, ref miss, ref miss,
                            ref replaceAll, ref miss, ref miss, ref miss, ref miss);
                    }

                    docs.Content.Font.Color = WdColor.wdColorBlack;

                    // TODO:possibly put this all on the top except saveas so we can use it when selecting the top intro part to acustom to job type

                    switch (output)
                    {
                        case "admin":
                            outputFI = @"M:\Troydon\Documents\Troydon\JobStuff\JobHunt\cover letters\Customs\\GeneralAdminCover.docx";
                            break;
                        case "itadmin":
                            outputFI = @"M:\Troydon\Documents\Troydon\JobStuff\JobHunt\cover letters\Customs\\ItAdminCover.docx";
                            break;
                        case "dev":
                            outputFI = @"M:\Troydon\Documents\Troydon\JobStuff\JobHunt\cover letters\Customs\\SoftWareDeveloperCover.docx";
                            break;
                        case "support":
                            outputFI = @"M:\Troydon\Documents\Troydon\JobStuff\JobHunt\cover letters\Customs\\ITSupportCover.docx";
                            break;
                        case "sales":
                            outputFI = @"M:\Troydon\Documents\Troydon\JobStuff\JobHunt\cover letters\Customs\\ITSalesCover.docx";
                            break;
                        default:
                            outputFI = @"M:\Troydon\Documents\Troydon\JobStuff\JobHunt\cover letters\Customs\\SoftWareDeveloperCover.docx";
                            break;
                    }
                    // keep looping until either the counter is not one of the readline values, and while the elected output is valid
                } while (outputMatch && (count == 0 || count == 1 || count == 2 || count == 3 || count == 4)); 

                // TODO: Next week: Check spelling, give spelling correction options if there are errors, and option to proceed without corrections, or option to finish after making modifitcation
                //var language = word.Languages[WdLanguageID.wdEnglishUS];
                //// Set the filename of the custom dictionary
                //// -- Based on:
                //// http://support.microsoft.com/kb/292108
                //// http://www.delphigroups.info/2/c2/261707.html
                //const string custDict = "custom.dic";

                //// Get the spelling suggestions
                //var suggestions = word.GetSpellingSuggestions("overfloww", custDict, MainDictionary: language.Name);

                //// Print each suggestion to the console
                //foreach (SpellingSuggestion spellingSuggestion in suggestions)
                //    Console.WriteLine("Suggested replacement: {0}", spellingSuggestion.Name);

                // TODO: make corrections


                docs.SaveAs(ref outputFI, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

                // then close the word document and close MS word process
            }
            //catch (COMException)
            //{
            //    // TODO: Add error messages
            //    CloseWordApplication(ref word, ref miss, ref docs);
            //}
            finally
            {
                CloseWordApplication(ref word, ref miss, ref docs);
            }

            //GC.Collect();
            //GC.WaitForPendingFinalizers();

            //GC.Collect();
            //GC.WaitForPendingFinalizers();
        }

        private static void ExitApp(ref Application word, ref object miss, ref Document docs, string output)
        {
            if (output == "exitapp")
            {
                Console.WriteLine("Application Shutting down");
                CloseWordApplication(ref word, ref miss, ref docs);
                Environment.Exit(0);
            }
        }

        // bool that returns false if any condition is true !CheckOutput(output) should be used to return true
        private static bool CheckOutput(string output)
        {
            if (output != "admin" && output != "itadmin" && output != "dev" && output != "support" && output != "sales") return true; // if none match return false
            else return false; // if one matches we return false and break the loop
        }
        private static void CloseWordApplication(ref Application word, ref object miss, ref Document docs)
        {
            if (docs != null)
            {
                docs.Close(
                    /* ref object SaveChanges */ ref miss,
                    /* ref object OriginalFormat */ ref miss,
                    /* ref object RouteDocument */ ref miss);
                docs = null;
            }

            if (word != null)
            {
                word.Quit(
                    /* ref object SaveChanges */ ref miss,
                    /* ref object OriginalFormat */ ref miss,
                    /* ref object RouteDocument */ ref miss);
                word = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private static void ExitApplication(ref Application word, ref object miss, ref Document docs, string appCommand)
        {
            if (appCommand == "exitapp")
            {
                Console.WriteLine("Application Shutting down");
                CloseWordApplication(ref word, ref miss, ref docs);
                Environment.Exit(0);
            }
        }
    }
}
