using Microsoft.Office.Interop.Word;
using System;
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

            // TODO: method to maybe increase length of what console.readline
            byte[] inputBuffer = new byte[1024];
            Stream inputStream = Console.OpenStandardInput(inputBuffer.Length);
            Console.SetIn(new StreamReader(inputStream, Console.InputEncoding, false, inputBuffer.Length));

            // Create new word application
            Application word = new Application();
            //word.Visible = true; // opens the doc
            // load all the MS Word data
            object miss = System.Reflection.Missing.Value;
            object path = @"M:\//troyi/Documents/Troydon/JobStuff/cover letters/Customs/SEEK_Free cover letter template_2018_NZ.docx";
            object readOnly = true;
            Document docs = word.Documents.Open(ref path, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

            // TODO: if possible execute the shutdown word method to execute if compilation shuts unexpectedly or cancelled to avoid it being left opened.
            //if (Environment.Exit(exitCode))

            try
            {
                //
                // determine outputfile name
                object outputFI;
                string output;
                do
                {
                        // Set user secrets for relevant skills 
                    string relevantSkills = "";
                    Console.WriteLine("What output would you like: 'admin' 'itadmin' 'dev' 'support' 'sales'");
                    output = Console.ReadLine();
                    if (output == "admin")
                        relevantSkills = "I have three years’ work experience working in a call centre as a Customer Service Rep with Woolworths Mobile. Some of the duties of the role included managing SAP tickets, whereby I had to administer events such as network incidents, transport and logistics, customer complaints, and input relevant data such as customer information, device details, funds, network incidents and faults. I was also known by the company to perform well with these duties.";
                    else if (output == "itadmin")
                        relevantSkills = "I have relevant skills useful to the role, I have achieved this when I completed a Certificate IV in proramming, this includes RDBMS such as SQL Server, MySQL, SQL, T-SQL. I have three years’ work experience working in a call centre as a Customer Service Rep with Woolworths Mobile. Some of the duties of the role included managing SAP tickets, whereby I had to administer events such as network incidents, transport and logistics, customer complaints, and input relevant data such as customer information, device details, funds, network incidents and faults. I was also known by the company to perform well with these duties.";
                    else if (output == "dev")
                        relevantSkills = "I have relevant skills useful to the role, which I obtained when I completed a Certificate IV in programming, that include basic C# OOP, ASP.NET Core, HTML, CSS, JavaScript, jQuery, SQL DBMS, Unity3D, Windows 10 OS, Microsoft Office 365 including Word and Excel, basic understanding of cloud services such as Azure and AWS. I am now applying these technologies into building my knowledge and skill level in creating beautiful, user friendly software, web and game applications. I am also working on using the programming, scripting and mark-up languages to learn the programming concepts such as algorithms, data structures, writing user-centric functional specifications, writing scalable code, understanding conditional logic, database design, responsive design.";
                    else if (output == "support")
                        relevantSkills = "I am interested in this position; I believe I have the skills and enthusiasm needed to do well. I am looking to secure a role that involves working with technology. I enjoy working with technology and dealing with computers, and Windows OS. I have knowledge in Office 364 Suite and great troubleshooting skills, which I have gained when working with Woolworths Mobile as a Tech Support Representative for mobile devices including devices such as Android, iOS and OPPO.";
                    else if (output == "sales")
                        relevantSkills = "I am interested in this position; I believe I have the skills and enthusiasm needed to do well. I am looking to secure a role that involves working with technology. I enjoy working with technology and dealing with computers, and Windows OS. I have knowledge in Office 364 Suite and great sales skills, which I have gained when working with Woolworths Mobile as a Tech Support Representative for mobile devices including devices such as Android, and OPPO."; 
                    // TODO: Maybe implement confirmation after each prompt later
                    // Gather all variables that we will need to insert into the document
                    string date = DateTime.Now.ToString("dddd, dd MMMM yyyy");

                    Console.WriteLine("Enter the Hiring Manger's Name (enter 'n' if none)");
                    string hiringManagersName = Console.ReadLine();
                    if (hiringManagersName == "n") hiringManagersName = ""; // if there is no name we reassign the value with an empty string

                    Console.WriteLine("Enter the company (hit 'n' if none)");
                    string company = Console.ReadLine();
                    if (company == "n") company = ""; // if there is no name we reassign the value with an empty string

                    string addressingEmployer = ""; // initial value will be Dear Sir/madam
                    if (hiringManagersName == "") addressingEmployer = "Sir/Madam"; // if response was  no, we will assign with the name in place of Sir/Madam
                    else if (hiringManagersName != "") addressingEmployer = hiringManagersName;

                    Console.WriteLine("Enter the role title");
                    string roleTitle = Console.ReadLine();

                    string advertiser = "Seek.com.au";
                    if (company != "n") advertiser = company; // if response was not no, we will assign with the name in place of <Company> 

                    Console.WriteLine("Is there a bonus question?, type 'n' if none");
                    string bonusQuestion = Console.ReadLine();
                    if (bonusQuestion == "n") bonusQuestion = "";

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

                    // TODO: Format everything to black font
                    docs.Content.Font.Color = WdColor.wdColorBlack;

                    // TODO:possibly put this all on the top except saveas so we can use it when selecting the top intro part to acustom to job type

                    switch (output)
                    {
                        case "admin":
                            outputFI = @"M:\troyi\Documents\Troydon\JobStuff\cover letters\Customs\\GeneralAdminCover.docx";
                            break;
                        case "itadmin":
                            outputFI = @"M:\troyi\Documents\Troydon\JobStuff\cover letters\Customs\\ItAdminCover.docx";
                            break;
                        case "dev":
                            outputFI = @"M:\troyi\Documents\Troydon\JobStuff\cover letters\Customs\\SoftWareDeveloperCover.docx";
                            break;
                        case "support":
                            outputFI = @"M:\troyi\Documents\Troydon\JobStuff\cover letters\Customs\\ITSupportCover.docx";
                            break;
                        case "sales":
                            outputFI = @"M:\troyi\Documents\Troydon\JobStuff\cover letters\Customs\\ITSalesCover.docx";
                            break;
                        default:
                            outputFI = @"M:\troyi\Documents\Troydon\JobStuff\cover letters\Customs\\SoftWareDeveloperCover.docx";
                            break;
                    }
                } while (output != "admin" && output != "itadmin" && output != "dev" && output != "support" && output != "sales"); // if all of these do not contain key word we keep looping

                // CHeck spelling
                var language = word.Languages[WdLanguageID.wdEnglishUS];

                // Set the filename of the custom dictionary
                // -- Based on:
                // http://support.microsoft.com/kb/292108
                // http://www.delphigroups.info/2/c2/261707.html
                const string custDict = "custom.dic";

                // Get the spelling suggestions
                var suggestions = word.GetSpellingSuggestions("overfloww", custDict, MainDictionary: language.Name);

                // Print each suggestion to the console
                foreach (SpellingSuggestion spellingSuggestion in suggestions)
                    Console.WriteLine("Suggested replacement: {0}", spellingSuggestion.Name);

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

            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();
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
        }
    }
}
