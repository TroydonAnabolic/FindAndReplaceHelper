using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.InteropServices;

namespace FindAndReplaceHelper
{
    class Program
    {
        static void Main(string[] args)
        { // TODO: Fix Word file being read not closing

            // Create new word application
            Application word = new Application();
            //word.Visible = true; // opens the doc
            // load all the MS Word data
            object miss = System.Reflection.Missing.Value;
            object path = @"M:\//troyi/Documents/Troydon/JobStuff/cover letters/Customs/SEEK_Free cover letter template_2018_NZ.docx";
            object readOnly = true;
            Document docs = word.Documents.Open(ref path, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

            try
            {
                // TODO: Maybe implement confirmation after each prompt later
                // Gather all variables that we will need to insert into the document
                string date = DateTime.Now.ToString("dddd, dd MMMM yyyy");

                Console.WriteLine("Enter the Hiring Manger's Name (enter 'no' if none)");
                string hiringManagersName = Console.ReadLine();
                if (hiringManagersName == "no") hiringManagersName = ""; // if there is no name we reassign the value with an empty string

                Console.WriteLine("Enter the company (hit 'no' if none)");
                string company = Console.ReadLine();
                if (company == "no") company = ""; // if there is no name we reassign the value with an empty string

                string addressingEmployer = ""; // initial value will be Dear Sir/madam
                if (hiringManagersName == "") addressingEmployer = "Sir/Madam"; // if response was  no, we will assign with the name in place of Sir/Madam
                else if (hiringManagersName != "")addressingEmployer = hiringManagersName; 

                Console.WriteLine("Enter the role title");
                string roleTitle = Console.ReadLine();

                string advertiser = "Seek.com.au";
                if (company != "no") advertiser = company; // if response was not no, we will assign with the name in place of <Company> 

                Console.WriteLine("Key-Skills. Here you want to highlight some of your core skills that talk to the key selection criteria for the role. E.g HTML");
                string keySkills = Console.ReadLine();

                Console.WriteLine("Talk about why you would like to work for the company and why you’d be a good fit. For example: Company name has been of interest" +
                    " to me since embarking on its mega store approach to retail. This is ideal for 21st century sales of flooring products. I was also impressed with" +
                    " the profile of your managing director Rod Smythe, which I read in the Retail journal late last year.");
                string whyWorkForCompany = Console.ReadLine();

                // Replace text now, loop through this 8 times replacing all the needed text
                for (int i = 0; i < 8; i++)
                {
                    Find contentReplace = word.Selection.Find;
                    contentReplace.ClearFormatting();

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
                            contentReplace.Text = "<Key-Skills. Here you want to highlight some of your core skills that talk to the key selection criteria for the role. E.g HTML>";
                            break;
                        case 7:
                            contentReplace.Text = "<Talk about why you would like to work for the company and why you’d be a good fit>";
                            break;
                        default:
                            break;
                    }

                    contentReplace.Replacement.ClearFormatting();
                    contentReplace.Replacement.Text = date;

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
                        case "<Key-Skills. Here you want to highlight some of your core skills that talk to the key selection criteria for the role. E.g HTML>":
                            contentReplace.Replacement.Text = keySkills;
                            break;
                        case "<Talk about why you would like to work for the company and why you’d be a good fit>":
                            contentReplace.Replacement.Text = whyWorkForCompany;
                            break;
                        default:
                            break;
                    }

                    object replaceAll = WdReplace.wdReplaceAll;
                    contentReplace.Execute(ref miss, ref miss, ref miss, ref miss, ref miss,
                        ref miss, ref miss, ref miss, ref miss, ref miss,
                        ref replaceAll, ref miss, ref miss, ref miss, ref miss);
                }

                // TODO: Format everything to black font
                docs.Content.Font.Color = WdColor.wdColorBlack;

                // determine outputfile name
                object outputFI;
                Console.WriteLine("What output would you like: 'admin' 'itadmin' 'dev' 'support' 'sales'");
                string output = Console.ReadLine();

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

                docs.SaveAs(ref outputFI, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

                // then close the word document and close MS word process
            }
            catch (COMException)
            {
            }
            finally
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

            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
