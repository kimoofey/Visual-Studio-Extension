namespace VSIXProject1
{
    using System.Diagnostics.CodeAnalysis;
    using System.Windows;
    using System.Windows.Controls;


    using System;
    using System.Text;
    using Microsoft.VisualStudio.Shell;
    using System.IO;
    using EnvDTE;

    /// <summary>
    /// Interaction logic for ToolWindow1Control.
    /// </summary>
    public partial class ToolWindow1Control : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ToolWindow1Control"/> class.
        /// </summary>
        public ToolWindow1Control()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// Handles click on the button by displaying a message box.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Justification = "Sample code")]
        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1300:ElementMustBeginWithUpperCaseLetter", Justification = "Default event handler naming pattern")]


        string result;
        string[] fileContent;

        private void Button1Click(object sender, RoutedEventArgs e)
        {
            result = "";
            TextBox1.Text = "";

            // Get DTE
            DTE dte = (DTE)Package.GetGlobalService(typeof(DTE));

            // Get Full Path to Document
            string docPath = dte.ActiveDocument.FullName;

            // Get File Content
            fileContent = File.ReadAllLines(path: Path.GetFullPath(dte.ActiveDocument.FullName));

            // Get File Code Model
            FileCodeModel fcm = dte.ActiveDocument.ProjectItem.FileCodeModel;
            CodeElements elts = null;
            elts = fcm.CodeElements;
            CodeElement elt = null;

            for (int i = 1; i <= fcm.CodeElements.Count; i++)
            {
                elt = elts.Item(i);
                SearchForFun(elt, elts, i);
            }
        }

        public void SearchForFun(CodeElement elt, CodeElements elts, long loc)
        {
            switch (elt.Kind)
            {
                case vsCMElement.vsCMElementNamespace:
                    {
                        CodeNamespace cns = null;
                        cns = (CodeNamespace)elt;
                        CodeElements memsVb = null;
                        memsVb = cns.Members;
                        for (int j = 1; j <= cns.Members.Count; j++)
                        {
                            SearchForFun(memsVb.Item(j), memsVb, j);
                        }
                        break;
                    }

                case vsCMElement.vsCMElementClass:
                    { 
                        //result += elem.Name + Environment.NewLine; break;
                        CodeClass cns = null;
                        cns = (CodeClass)elt;
                        CodeElements memsVb = null;
                        memsVb = cns.Members;
                        for (int j = 1; j <= cns.Members.Count; j++)
                        {
                            SearchForFun(memsVb.Item(j), memsVb, j);
                        }
                        break;
                    }

                case vsCMElement.vsCMElementFunction:
                    {
                        int startLine = elt.StartPoint.Line;
                        int endLine = elt.EndPoint.Line;
                        int linesEmpty = 0;
                        int linesComments = 0;
                        int amountKeyWords = 0;
                        string test = "";
                        FunctionHandler(startLine, endLine, ref linesEmpty, ref linesComments, ref amountKeyWords, ref test);
                        result += elt.Name + Environment.NewLine +"Total Lines: " + (endLine - startLine + 1).ToString() 
                            + Environment.NewLine + "Empty Lines: " + linesEmpty.ToString() + Environment.NewLine + 
                            "Comment Lines: " + linesComments.ToString() + Environment.NewLine + "Key Words: " 
                            + amountKeyWords.ToString() + Environment.NewLine + Environment.NewLine;
                        break;
                    }

                default: break;
            }

            TextBox1.Text = result;
        }

        private void FunctionHandler(int startLine, int endLine, ref int linesEmpty,
            ref int lineComments, ref int amountKeyWords, ref string testString)
        {
            // Calculate Empty Lines
            for (int i = startLine; i < endLine; i++)
            {
                if (fileContent[i - 1].Length == 0)
                {
                    linesEmpty++;
                }
            }

            // Searching for everything
            bool isString = false;
            bool isComment = false;
            bool isBigComment = false;
            for (int i = startLine; i < endLine; i++)
            {
                StringBuilder sb = new StringBuilder(fileContent[i - 1]);
                bool isLineCommented = isComment || isBigComment;
                if (isComment == true)
                {
                    for (int j = 0; j < sb.Length; ++j)
                    {
                        sb[j] = 'A';
                    }

                    if (fileContent[i - 1].Length > 0 && fileContent[i - 1][fileContent[i - 1].Length - 1] == '\\')
                    { }
                    else
                    {
                        isComment = false;
                        // uncomment for not to calculate comments lines if they are empty
                        // continue;
                    }
                }

                for (int j = 0; j < fileContent[i - 1].Length; ++j)
                {
                    if (isString == true)
                    {
                        if (fileContent[i - 1][j] == '\"' && j - 1 >= 0 && fileContent[i - 1][j - 1] != '\\')
                        {
                            isString = false;
                            continue;
                        }

                        sb[j] = 'A';
                    }
                    else if (isBigComment == true)
                    {
                        if (fileContent[i - 1][j] == '*' && j + 1 < fileContent[i - 1].Length && fileContent[i - 1][j + 1] == '/')
                        {
                            isBigComment = false;
                            continue;
                        }

                        sb[j] = 'A';
                    }
                    else
                    {
                        if (fileContent[i - 1][j] == '\"')
                        {
                            isString = true;
                        }
                        else if (fileContent[i - 1][j] == '/' && j + 1 < fileContent[i - 1].Length && fileContent[i - 1][j + 1] == '/')
                        {
                            isLineCommented = true;
                            if (fileContent[i - 1][fileContent[i - 1].Length - 1] == '\\')
                            {
                                isComment = true;
                            }

                            for (; j < sb.Length; ++j)
                            {
                                sb[j] = 'A';
                            }

                            break;
                        }
                        else if (fileContent[i - 1][j] == '/' && j + 1 < fileContent[i - 1].Length && fileContent[i - 1][j + 1] == '*')
                        {
                            isBigComment = true;
                            isLineCommented = true;
                        }
                    }
                }

                if (isLineCommented)
                {
                    lineComments++;
                }

                fileContent[i - 1] = sb.ToString();
            }

            // Calculating key words
            CalculateKeyWords(startLine, endLine, ref amountKeyWords);
        }

        private void CalculateKeyWords(int startLine, int endLine, ref int amountKeyWords)
        {
            string[] arr = {"and", "and_eq", "asm", "auto","bitor", "bool", "break", "case",
                "catch", "char","class", "const", "constexpr", "const_cast","continue",
                "decltype", "default", "delete", "do", "double", "dynamic_cast","else",
                "enum", "explicit", "export", "extern", "false", "float", "for", "friend",
                "goto", "if", "inline", "int", "long", "mutable", "namespace", "new",
                "noexcept", "not", "nullptr", "operator", "or", "private", "protected",
                "public", "register", "reinterpret_cast", "return", "short", "signed",
                "sizeof", "static", "static_assert", "static_cast","struct", "switch",
                "template", "this", "thread_local", "throw", "true", "try", "typedef",
                "typeid", "typename", "union", "unsigned", "using", "virtual", "void",
                "volatile", "while", "xor"
            };

            for (int i = startLine; i <= endLine; i++)
            {
                for (int k = 0; k < arr.Length; k++)
                {
                    string key = arr[k];
                    
                    //StringBuilder sb = new StringBuilder(fileContent[i]);
                    string str = fileContent[i - 1];
                    while (true)
                    {
                        if (str.IndexOf(key) == -1)
                        {
                            break;
                        }

                        if ((str.IndexOf(key) - 1 >= 0) && str[str.IndexOf(key) - 1] >= 'A' && str[str.IndexOf(key) - 1] <= 'z')
                        {
                            str = str.Substring(str.IndexOf(key) + key.Length);
                            continue;
                        }

                        int idxPlus = str.IndexOf(key) + key.Length;
                        if ((idxPlus < str.Length) && str[idxPlus] >= 'A' && str[idxPlus] <= 'z')
                        {
                            str = str.Substring(str.IndexOf(key) + key.Length);
                            continue;
                        }

                        str = str.Substring(str.IndexOf(key) + key.Length);
                        amountKeyWords++;
                    }
                }
            }
        }

        private void TextBoxTextChanged(object sender, TextChangedEventArgs e)
        {
            //MessageBox.Show("Done");
        }
    }
}