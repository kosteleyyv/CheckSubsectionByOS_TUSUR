using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace CheckSubsectionByOS_TUSUR
{
    internal class DocumentCheckUp
    {
        public static class ObjectTitleStrings
        {
            public static String FigureTitle { get { return "Рисунок"; } }
            public static String TableTitle { get { return "Таблица"; } }
            public static String CodeTitle { get { return "Листинг"; } }

            public static String FigureRef { get { return "рис"; } }
            public static String TableRef { get { return "табл"; } }
            public static String CodeRef { get { return "лист"; } }
        }

        private class ParagraphInfo
        {
            public int Index = 0;
            public enum ParagraphClass
            {
                Empty,
                Текст,
                Код,
                Рисунок,
                Таблица,
                Заголовок,
                ПодрисуночнаяПодпись,
                НазваниеТаблицы,
                НазваниеЛистинга,
                ЭлементНумерованногоСписка,
                ЭлементМаркерованногоСписка,
                БиблиографическоеОписаниеИсточника,
                ЗаголовокСпискаЛитературы,
                NumberText,
                NumberList
            };

            public ParagraphClass Type = ParagraphClass.Empty;

            /// текст примечаний в документе
            public List<string> Problems = new List<string>();

            /// номер таблицы, рисунка или листинга
            public int IndexObject = 0;
            /// номер таблицы, рисунка или листинга в тексте работы
            public string NumberObjectInText = null;
            /// есть ли ссылка на таблицу, рисунок или листинг в предыдущих абзацах
            public bool HasRef = true;
            /// абзац - последний элемент списка
            public bool isLastListElement = false;
            /// абзац - текст перед списком
            public bool isTextBeforeList = false;
            /// пустая ячейка с символом \a
            public bool isEmptyTableCellWithA = false;
        }

        private class DocumentInfo
        {
            public bool HasTitle = false;
            public bool HasSource = false;
            public bool HasReference = false;

            public bool HasGeneralComments = false;
        }

        public static readonly string[] OperatorsList = new string[]
    {
        // Арифметические операторы
        "+", "-", "*", "/", "%",
        "++", "--",
        
        // Операторы присваивания
        "=", "+=", "-=", "*=", "/=", "%=",
        "&=", "|=", "^=", "<<=", ">>=",
        
        // Операторы сравнения
        "==", "!=", ">", "<", ">=", "<=",
        
        // Логические операторы
        "&&", "||", "!",
        
        // Побитовые операторы
        "&", "|", "^", "~", "<<", ">>",
        
        // Операторы указателей и адресов
        "&", "*", "->", ".*", "->*",
        
        // Операторы вызова и доступа
       /* "()", "[]", ".",*/ "::",
        
        // Тернарный оператор
        "?:",

        "}","{"
        
        //// Операторы управления памятью
        //"new", "delete", "new[]", "delete[]",
        
        //// Операторы приведения типов
        //"static_cast", "dynamic_cast", "const_cast", "reinterpret_cast",
        
        //// Прочие операторы
        //",", "sizeof", "typeid", "noexcept"
    };

        static readonly string[] KeywordsList = new string[] {
    // Основные
  
    "auto",
    "bool",
    "break",
    "case",
    "catch",
    "char",
    "char8_t",
    "char16_t",
    "char32_t",
    "class",
    "const",
    "constexpr",
    "const_cast",
    "continue",
    "decltype",
    "default",
    "delete",
    "do",
    "double",
    "dynamic_cast",
    "else",
    "enum",
    "explicit",
    "export",
    "extern",
    "false",
    "float",
    "for",
    "friend",
    "if",
    "inline",
    "int",
    "long",
    "namespace",
    "new",
    "noexcept",
    "nullptr",
    "operator",
    "or",
    "private",
    "protected",
    "public",
    "reinterpret_cast",
    "return",
    "short",
    "signed",
    "sizeof",
    "static",
    "static_assert",
    "static_cast",
    "struct",
    "switch",
    "template",
    "this",
    "throw",
    "true",
    "try",
    "typedef",
    "typeid",
    "typename",
    "union",
    "unsigned",
    "using",
    "virtual",
    "void",
    "volatile",
    "wchar_t",
    "while"
       };


        static readonly string ListMarkers = "\\u2022\\u25aa\\u006f\\u2014\\u2013\\u202d";
        // \\u2022 круглый маркер •
        // \\u25aa квадратный маркер ▪
        // \\u006f o
        // \\u2014 длинное тире
        // \\u2013 короткое тире
        // \\u202d ???

        static readonly string OSListMarkers = "\\u2022\\u25aa\\u2013";

        private static int ContainsKeyWordOrOperator(string word)
        {
            int count = 0;

            for (int i = 0; i < OperatorsList.Length; i++)
            {
                if (word.Contains(OperatorsList[i]))
                {
                    count++;
                }
            }

            for (int i = 0; i < KeywordsList.Length; i++)
            {
                if (word.Contains(KeywordsList[i]))
                {
                    count++;
                    // break; ???
                }
            }

            return count;
        }

        private static bool isCyrilic(Microsoft.Office.Interop.Word.Range range, out bool isCouriewNew)
        {
            var regex = new System.Text.RegularExpressions.Regex("[а-я]");
            int cyrrilicCount = 0;
            int alphaWordCount = 0;
            var alpha = new Regex("[a-zа-я]");
            int courierNewWordCount = 0;

            for (int i = 1; i <= range.Words.Count; i++)
            {
                string word = range.Words[i].Text.Trim().ToLower();

                if (alpha.IsMatch(word)) // считаем,сколько букв
                {
                    int n = regex.Matches(word).Count;
                    if (n > 0.6 * word.Length) // если более 60 % - кириллическое слово
                    {
                        cyrrilicCount++;
                    }
                    // var style = range.Words[i].CharacterStyle as Style;

                    if (range.Words[i].Font.Name == "Courier New")
                    {
                        courierNewWordCount++;
                    }

                    alphaWordCount++;
                }


            }
            isCouriewNew = (courierNewWordCount * 100.0 / alphaWordCount > 90.0);

            // TODO проверь ключевые слова и операторы

            return (100.0 * cyrrilicCount / alphaWordCount > 60);
        }

        /// извлекает первое "слово" из диапазоно, в т.ч. маркеры и цифры
        private static string ExtractFirstWord(Microsoft.Office.Interop.Word.Range range)
        {
            var regex = new System.Text.RegularExpressions.Regex($"[1-9a-zа-я{ListMarkers}]");

            for (int i = 1; i <= range.Words.Count; i++)
            {
                string word = range.Words[i].Text.Trim().ToLower();

                if (word.Length != 0 && regex.IsMatch(word[0] + ""))
                {
                    return word;
                }
            }
            return null;
        }

        /// проверка библиографического описания
        private static void checkSource(Paragraph paragraph, ParagraphInfo paragraphInfo, DocumentInfo documentParams)
        {
            var paragraphRange = paragraph.Range;
            string text = paragraphRange.Text;
            text = text.Trim();

            if (text.Length != 0 && text[text.Length - 1] != '.')
            {
                paragraphInfo.Problems.Add("должна быть точка в конце абзаца");
            }

            string[] denyURLs = new string[] { "wikipedia.org", "habr.com" };
            bool hasDenyURLs = false;

            for (int i = 0; i < denyURLs.Length; i++)
            {
                var range = paragraph.Range;
                var find = range.Find;
                find.ClearFormatting();
                find.Text = denyURLs[i];

                while (find.Execute())
                {
                    if (range.End > paragraph.Range.End)
                    {
                        break;
                    }

                    hasDenyURLs = true;
                    range.HighlightColorIndex = WdColorIndex.wdYellow;
                    range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                }
            }

            if (hasDenyURLs)
            {
                paragraphInfo.Problems.Add("нельзя использовать Википедию или Хабр в качестве источников, как и другие noname-сайты. " +
                                           "Используем учебники, официальную документацию или другие источники, написанные признанными специалистами.");
            }
        }

        // выделить цветом совпадения
        static void Highlight(Paragraph paragraph, MatchCollection matches, WdColorIndex color = WdColorIndex.wdYellow)
        {
            for (int i = 0; i < matches.Count; i++)
            {
                var range = paragraph.Range;
                var find = range.Find;
                find.ClearFormatting();
                find.Text = matches[i].Value;

                while (find.Execute())
                {
                    if (range.End > paragraph.Range.End)
                    {
                        break;
                    }
                    range.HighlightColorIndex = color;
                    range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                }
            }
        }
        /// проверка абзаца обычного текста в т.ч. списка
        private static void checkText(Paragraph paragraph, ParagraphInfo paragraphInfo, DocumentInfo documentParams)
        {

            string text = paragraph.Range.Text;

            // с большой буквы и на конце точка
            bool isV1 = false;
            // маркерованный или иной список
            bool isV2 = false;
            // с большой буквы и на конце двоеточие
            bool isV3 = false;

            if (paragraphInfo.Type == ParagraphInfo.ParagraphClass.ЭлементНумерованногоСписка)
            {
                // для параграфа определен нумерованный список
                if (paragraph.Range.ListParagraphs.Count != 0)
                {
                    // извлекаем маркер и основной текст
                    string marker = paragraph.Range.ListFormat.ListString.Trim();
                    text = text.Trim();

                    // маркер с точкой
                    if (marker.EndsWith("."))
                    {
                        isV1 = true;
                    }
                    else
                    {
                        isV2 = true;
                    }
                }
                else
                {
                    // список не определен студентом, а номер введен вручную
                    // извлекаем текст
                    text = text.Trim();

                    var regexNumber = new Regex("[1-9][0-9]*[.)]");

                    if (regexNumber.IsMatch(text))
                    {
                        string marker = regexNumber.Match(text).Value;
                        if (marker.EndsWith("."))
                        {
                            isV1 = true;
                        }
                        else
                        {
                            isV2 = true;
                        }
                    }

                    // надо вручную проверить наличие пробела
                    var regexNumberWithWS = new Regex("[1-9][0-9]*[.)][\\s]");
                    if (regexNumberWithWS.IsMatch(text))
                    {
                        text = regexNumberWithWS.Replace(text, "");
                    }
                    else
                    {
                        text = regexNumber.Replace(text, "");
                        paragraphInfo.Problems.Add("между номером и предложением должен быть отступ в виде пробела или табуляции");
                    }

                }
            }

            if (paragraphInfo.Type == ParagraphInfo.ParagraphClass.ЭлементМаркерованногоСписка)
            {
                isV2 = true;

                if (paragraph.Range.ListParagraphs.Count == 0)
                {
                    text = text.Trim();

                    var regexMarker = new Regex($"[{ListMarkers}]");

                    var regexMarkerWithWS = new Regex($"[{ListMarkers}][\\s]");

                    var regexMarkerAdvance = new Regex($"[{OSListMarkers}]");

                    if (regexMarkerWithWS.IsMatch(text))
                    {
                        paragraphInfo.Problems.Add("рекомендуемый тип маркера списка: тире(–), точка(•), квадрат(▪)");
                    }

                    if (regexMarkerWithWS.IsMatch(text))
                    {
                        text = regexMarkerWithWS.Replace(text, "");
                    }
                    else
                    {
                        text = regexMarker.Replace(text, "");
                        paragraphInfo.Problems.Add("между маркером и предложением должен быть отступ в виде пробела или табуляции");
                    }


                }
            }

            // обычный текст
            if (paragraphInfo.Type == ParagraphInfo.ParagraphClass.Текст)
            {
                if (paragraphInfo.isTextBeforeList)
                {
                    isV3 = true;
                }
                else
                {
                    isV1 = true;
                }

                if ((text.StartsWith("\t") || text.StartsWith(" ")))
                {
                    paragraphInfo.Problems.Add("убрать пробел или табуляцию в начале предложения");
                }
            }

            // получили текст без маркера или номера
            text = text.Trim();

            if (isV1)
            {
                if (text.Length != 0 && !Char.IsUpper(text[0]))
                {
                    paragraphInfo.Problems.Add("абзац должен начинаться с большой буквы");
                }

                if (text.Length != 0 && text[text.Length - 1] != '.')
                {
                    paragraphInfo.Problems.Add("должна быть точка в конце абзаца");
                }
            }

            if (isV2)
            {
                if (text.Length != 0 && Char.IsUpper(text[0]))
                {
                    paragraphInfo.Problems.Add("абзац должен начинаться с маленькой буквы");
                }

                if (paragraphInfo.isLastListElement && text.Length != 0 && text[text.Length - 1] != '.')
                {
                    paragraphInfo.Problems.Add("должна быть точка в конце абзаца");
                }

                if (!paragraphInfo.isLastListElement && text.Length != 0 && text[text.Length - 1] != ';')
                {
                    paragraphInfo.Problems.Add("должна быть точка с запятой в конце абзаца");
                }

                // TODO сделать выделение последнего символа и первого символа
            }

            if (isV3)
            {
                if (text.Length != 0 && !Char.IsUpper(text[0]))
                {
                    paragraphInfo.Problems.Add("абзац должен начинаться с большой буквы");
                }

                if (text.Length != 0 && text[text.Length - 1] != ':')
                {
                    paragraphInfo.Problems.Add("должно быть двоеточие в конце абзаца (перед списком)");
                }
            }

            // слова я или мы или вы
            var regex = new Regex("([^a-zA-Zа-яА-Я]|^)((я)|(Я)|(мы)|(Мы)|(вы)|(Вы)|(нам)|(вам)|(Нам)|(Вам))[^a-zA-Zа-яА-Я]");

            if (regex.IsMatch(text))
            {
                paragraphInfo.Problems.Add("пишем обезличенно без я, мы, вы");
                var matches = regex.Matches(text);
                Highlight(paragraph, matches);
            }

            // применили дефис не между двумя словами
            regex = new Regex("([^a-zA-Zа-яА-Я])(-|—)[^a-zA-Zа-яА-Я>]");
            if (regex.IsMatch(text))
            {
                paragraphInfo.Problems.Add("использовать правильно тире – вместо дефиса - и длинного тире —");
                var matches = regex.Matches(text);
                Highlight(paragraph, matches);
            }

            regex = new Regex("\\[[0-9]{1,3}\\]"); // TODO изменить регулярку без учетам 01 или 001
            if (regex.IsMatch(text))
            {
                documentParams.HasReference = true;

                Regex[] variants = new[] {
                    new Regex("[\\S]\\[[0-9]{1,3}\\]"),     // между ссылкой и словом нет пробела
                    new Regex("[.][\\s]?\\[[0-9]{1,3}\\]"), // перед ссылкой точка
                    new Regex("[\\u00A0]\\[[0-9]{1,3}\\]"),
                    new Regex("[^\\u00A0]\\[[0-9]{1,3}\\]") // д.б. неразрывный
                };

                if (variants[0].IsMatch(text)) // между ссылкой и словом нет пробела
                {
                    var matches = variants[0].Matches(text);
                    paragraphInfo.Problems.Add("между ссылкой на источник и словом должен быть неразрывный пробел");
                    Highlight(paragraph, matches);
                }
                else
                {   // используемый пробел не неразрывный!
                    if (!variants[2].IsMatch(text))
                    {
                        var matches = variants[3].Matches(text);
                        paragraphInfo.Problems.Add("между ссылкой на источник и словом должен быть неразрывный пробел (shift+ctrl+пробел)");
                        Highlight(paragraph, matches);
                    }
                }

                if (variants[1].IsMatch(text))
                {
                    var matches = variants[1].Matches(text);
                    paragraphInfo.Problems.Add("ссылка на источник входит в предложение, поэтому точка ставится после ссылки");
                    Highlight(paragraph, matches);
                }
            }

            regex = new Regex("[\"“‟”„''‚‘’`]");
            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);
                paragraphInfo.Problems.Add("необходимо использовать «кавычки-ёлочки», а не кавычки-лапки и одиночные кавычки");
                Highlight(paragraph, matches);
            }

            regex = new Regex("[^«][a-zA-Zа-яА-Я]*_[a-zA-Zа-яА-Я]*([(][)])?[^»]");

            if (regex.IsMatch(text))
            {
                if (!new Regex("[«][a-zA-Zа-яА-Я]*_[a-zA-Zа-яА-Я]*([(][)])[»]").IsMatch(text))
                {
                    // нашли слово, которое следует обернуть в кавычки
                    var matches = regex.Matches(text);
                    paragraphInfo.Problems.Add("латинские названия или названия с _ следует обернуть в кавычки-елочки");
                    Highlight(paragraph, matches);
                }
            }

            regex = new Regex("[^«][a-zA-Zа-яА-Я]+::[a-zA-Zа-яА-Я]+([(][)])?[^»]");

            if (regex.IsMatch(text))
            {
                if (!new Regex("[«][a-zA-Zа-яА-Я]+::[a-zA-Zа-яА-Я]+([(][)])[»]").IsMatch(text))
                {
                    var matches = regex.Matches(text);
                    paragraphInfo.Problems.Add("латинские названия с пространством имен следует обернуть в кавычки-елочки");
                    Highlight(paragraph, matches);
                }
            }

            regex = new Regex("[(][)]");

            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);
                paragraphInfo.Problems.Add("не используем пустые () и не указываем () в названиях функций или методов");
                Highlight(paragraph, matches);
            }

            regex = new Regex("[^«]([+*&^$~><=]|([\\+]{2})|([-]{2})|(->)|(>>)|(<<)|([\\+]=)|(-=)|([\\*]=)|(\\/=)|(>=)|(<=)|(!=)|(&&)|(::)|([\\|]{2}))[^»]");

            if (regex.IsMatch(text))
            {
                bool isNotCppName = false;
                var matches = regex.Matches(text);

                for (int i = 0; i < matches.Count; i++)
                {
                    if (matches[i].Value.ToLower().StartsWith("c+") || matches[i].Value.ToLower().StartsWith("с+"))
                    {
                        continue;
                    }

                    isNotCppName = true;

                    var range = paragraph.Range;
                    var find = range.Find;
                    find.ClearFormatting();
                    find.Text = matches[i].Value;

                    while (find.Execute())
                    {
                        if (range.End > paragraph.Range.End)
                        {
                            break;
                        }
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }

                if (isNotCppName)
                {
                    paragraphInfo.Problems.Add("знаки операторов следует обернуть в кавычки-елочки");
                }
            }

            regex = new Regex("[A-ZА-Я]{2,}");

            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);
                paragraphInfo.Problems.Add("напоминание: не забывайте расшифровать аббревиатуру в месте ее первого использования, например, программное обеспечение (ПО)");
                Highlight(paragraph, matches, WdColorIndex.wdGray50);
            }

            regex = new Regex("([A-ZА-Я]+[^a-zA-Zа-яА-Я]){3,}");
            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);
                paragraphInfo.Problems.Add("не должно быть капсола в тексте (или у Вас три подряд аббревиатуры - так можно)");
                Highlight(paragraph, matches);
            }

            regex = new Regex(".[\\s]{2,}.");
            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);
                paragraphInfo.Problems.Add("обнаружен множественный пробел, нужно сократить до одного");
                Highlight(paragraph, matches);
            }

            bool denyHyperlinks = false;

            if (paragraph.Range.Hyperlinks.Count != 0)
            {
                for (int i = 1; i <= paragraph.Range.Hyperlinks.Count; i++)
                {
                    string url = paragraph.Range.Hyperlinks[i].Address;

                    if (url != null) // ссылка на сайт - нельзя
                    {
                        paragraph.Range.Hyperlinks[i].Range.HighlightColorIndex = WdColorIndex.wdYellow;
                        denyHyperlinks = true;
                    }
                }
            }

            if (denyHyperlinks)
            {
                paragraphInfo.Problems.Add("убрать гиперссылки на интернет-страницы из текста." +
                    "Прим. Ссылкой на источник является текст вида [1], который может определять только перекрестную ссылку на элемент списка литературы (номер списка), а не переход по URL-ссылки");
            }


            regex = new Regex("\\t");
            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);
                paragraphInfo.Problems.Add("убрать знак табуляции внутри текста");
                for (int i = 0; i < matches.Count; i++)
                {
                    var range = paragraph.Range;
                    var find = range.Find;
                    find.ClearFormatting();
                    find.Text = matches[i].Value;

                    while (find.Execute())
                    {
                        if (range.End > paragraph.Range.End)
                        {
                            break;
                        }

                        var newRange = paragraph.Range; // выделяем, задевая боковые буквы
                        newRange.Start = range.Start - 1;
                        newRange.End = range.End + 1;

                        newRange.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }
        }

        /// проверка заголовка
        static string checkHeader(Paragraph paragraph, ParagraphInfo paragraphInfo, DocumentInfo documentParams)
        {
            string levelNumber = "";

            string text = paragraph.Range.Text.Trim();

            if (paragraph.Range.ListParagraphs.Count != 0)
            {
                string marker = paragraph.Range.ListFormat.ListString.Trim();

                if (marker.EndsWith("."))
                {
                    paragraphInfo.Problems.Add("не должно быть точки на конце номера раздела или подраздела");
                }

                levelNumber = marker.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0];
            }
            else
            {
                string headerText = paragraph.Range.Text.Trim();
                var regexNumber = new Regex("[1-9]([.][1-9][0-9]*)*.");

                string marker = regexNumber.Match(headerText).Value;

                if (marker.EndsWith("."))
                {
                    paragraphInfo.Problems.Add("не должно быть точки на конце номера раздела или подраздела");

                    regexNumber = new Regex("[1-9]([.][1-9][0-9]*)*[.].");
                    marker = regexNumber.Match(headerText).Value;
                }

                if (!new Regex("[\\s]").IsMatch(marker[marker.Length - 1] + ""))
                {
                    paragraphInfo.Problems.Add("между номером заголовка и текстом заголовка должен быть пробел или табуляция");
                }

                marker = marker.Trim();

                text = new Regex("[1-9]([.][1-9][0-9]*)*[.]?[\\s]?").Replace(text, "");

                levelNumber = marker.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0];
            }

            var regex = new Regex(".[\\s]{2,}.");

            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);
                paragraphInfo.Problems.Add("обнаружен множественный пробел, нужно сократить до одного");
                Highlight(paragraph, matches);
            }

            regex = new Regex("([^a-zA-Zа-яА-Я])(-|—)[^a-zA-Zа-яА-Я]");

            if (regex.IsMatch(text))
            {
                paragraphInfo.Problems.Add("использовать правильно тире – вместо дефиса - и длинного тире —");
                var matches = regex.Matches(text);
                Highlight(paragraph, matches);
            }

            regex = new Regex("[\"“‟”„''‚‘’`]");
            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);
                paragraphInfo.Problems.Add("необходимо использовать «кавычки-ёлочки», а не кавычки-лапки и одиночные кавычки");
                Highlight(paragraph, matches);
            }

            if (text.EndsWith("."))
            {
                paragraphInfo.Problems.Add("не должно быть точки на конце");
            }

            if (text.Length > 0 && !Char.IsUpper(text[0]))
            {
                paragraphInfo.Problems.Add("текст должен быть с заглавной буквы");
            }

            return levelNumber;
        }


        static string checkObjectTitle(Paragraph paragraph, ParagraphInfo paragraphInfo, DocumentInfo documentParams, string marker, string level, ref int number)
        {
            string text = paragraph.Range.Text.Trim();

            var regex = new Regex(".[\\s]{2,}.");

            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);
                paragraphInfo.Problems.Add("обнаружен множественный пробел, нужно сократить до одного");
                Highlight(paragraph, matches);
            }

            regex = new Regex("([^a-zA-Zа-яА-Я])(-|—)[^a-zA-Zа-яА-Я]");
            if (regex.IsMatch(text))
            {
                paragraphInfo.Problems.Add("использовать правильно тире – вместо дефиса - и длинного тире —");
                var matches = regex.Matches(text);
                Highlight(paragraph, matches);
            }

            regex = new Regex("[\"“‟”„''‚‘’`]");
            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);
                paragraphInfo.Problems.Add("необходимо использовать «кавычки-ёлочки», а не кавычки-лапки и одиночные кавычки");
                Highlight(paragraph, matches);
            }

            if (text.EndsWith("."))
            {
                paragraphInfo.Problems.Add("не должно быть точки на конце");
            }

            if (text.Length > 0 && !Char.IsUpper(text[0]))
            {
                paragraphInfo.Problems.Add("Текст должен быть с заглавной буквы");
            }

            regex = new Regex(marker + "[\\s][1-9][\\d]*[.][1-9][\\d]*[\\s]–[\\s][A-ZА-Я]");

            if (!regex.IsMatch(text))
            {
                paragraphInfo.Problems.Add($"корректный формат начала подписи объекта: {marker} {level}.{number} – Текст подписи с заглавной буквы");
            }

            regex = new Regex("[1-9][\\d]*[.][1-9][\\d]*");

            if (regex.IsMatch(text))
            {
                var match = regex.Match(text);

                if (level != null) // знаем номер раздела
                {
                    if (match.Value != $"{level}.{number}")
                    {
                        bool isHighlight = false;
                        if (number == 1)
                        {
                            var strings = match.Value.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                            if (strings.Length >= 2)
                            {
                                int.TryParse(strings[1], out number);
                            }

                            if (match.Value != $"{level}.{number}")
                            {
                                paragraphInfo.Problems.Add($"не совпадает номер объекта с ожидаемым: {marker} {level}.{number}");
                                isHighlight = true;
                            }
                        }
                        else
                        {
                            paragraphInfo.Problems.Add($"не совпадает номер объекта с ожидаемым: {marker} {level}.{number}");
                            isHighlight = true;
                        }

                        if (isHighlight == true)
                        {
                            var range = paragraph.Range;
                            var find = range.Find;
                            find.ClearFormatting();
                            find.Text = match.Value;

                            while (find.Execute())
                            {
                                if (range.End > paragraph.Range.End)
                                {
                                    break;
                                }
                                range.HighlightColorIndex = WdColorIndex.wdYellow;
                                range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                            }
                        }

                    }
                }
                else
                {
                    paragraphInfo.Problems.Add($"не проверена нумерация из-за отсутствия номер раздела");
                }

                return match.Value;
            }
            else
            {
                regex = new Regex("[1-9][\\d]*");

                paragraphInfo.Problems.Add($"номер объекта должен совпадать с форматом {marker} {level}.{number} (номер раздела.номер объекта");

                if (regex.IsMatch(text))
                {
                    var match = regex.Match(text);

                    var range = paragraph.Range;
                    var find = range.Find;
                    find.ClearFormatting();
                    find.Text = match.Value;

                    while (find.Execute())
                    {
                        if (range.End > paragraph.Range.End)
                        {
                            break;
                        }
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }

                    return match.Value;
                }

            }

            return null;
        }

        public static void checkDocument(string pathDoc)
        {
            List<ParagraphInfo> infos = new List<ParagraphInfo>();
            DocumentInfo documentParams = new DocumentInfo();

            var application = new Microsoft.Office.Interop.Word.Application();
            application.Visible = true;

            Document document = application.Documents.Open(pathDoc, false);

            // удаляем комменты прошлого запуска
            for (int i = document.Comments.Count; i >= 1; i--)
            {
                if (document.Comments[i].Author == "ROBOT")
                {
                    if (document.Comments[i].Replies.Count == 0)
                    {
                        document.Comments[i].DeleteRecursively();
                    }
                }
            }

            int paragraphIndex = 0;
            // определение типа параграфа в соответствии с ParagraphInfo.ParagraphType
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                //application.Selection.SetRange(paragraph.Range.Start, paragraph.Range.End);

                ParagraphInfo paragraphInfo = new ParagraphInfo();
                paragraphInfo.Index = ++paragraphIndex;

                // анализ текста
                string text = paragraph.Range.Text.Trim().ToLower();

                if (paragraph.Range.InlineShapes.Count != 0) // есть рисунок
                {
                    // TODO если поставят один символ случайно, то надо сообщить???
                    paragraphInfo.Type = ParagraphInfo.ParagraphClass.Рисунок;
                    infos.Add(paragraphInfo);
                    continue;
                }

                // пустая строка - м.б. отступом в коде или рисунком
                if (text.Length == 0)
                {
                    infos.Add(paragraphInfo);
                    continue;
                }

                if (documentParams.HasSource) // ниже заголовка списка источников находится их перечисление
                {
                    paragraphInfo.Type = ParagraphInfo.ParagraphClass.БиблиографическоеОписаниеИсточника;
                    infos.Add(paragraphInfo);
                    continue;
                }

                // если есть таблица - или таблица или код в рамке (в старых работах)
                if (paragraph.Range.Tables.Count != 0)
                {
                    string textTable = paragraph.Range.Text.Replace("\a", "").Trim();

                    if (textTable.Length == 0)
                    {
                        paragraphInfo.Type = ParagraphInfo.ParagraphClass.Таблица;
                        paragraphInfo.isEmptyTableCellWithA = true;
                        infos.Add(paragraphInfo);
                        continue;
                    }
                    bool isCourierNew = false;
                    // проверить длину текста
                    if (!isCyrilic(paragraph.Range, out isCourierNew) || isCourierNew)
                    {
                        paragraphInfo.Type = ParagraphInfo.ParagraphClass.Код;
                        infos.Add(paragraphInfo);
                        continue;
                    }

                    paragraphInfo.Type = ParagraphInfo.ParagraphClass.Таблица;
                    infos.Add(paragraphInfo);
                    continue;
                }

                // регулярки на определение маркированного или нумерованного текста
                var regexNumber = new System.Text.RegularExpressions.Regex("[1-9]");
                var regexMarker = new System.Text.RegularExpressions.Regex($"[{ListMarkers}]");

                if (paragraph.Range.ListParagraphs.Count != 0) // если есть список
                {
                    switch (paragraph.Range.ListFormat.ListType) // разделение на маркерованный и нумерованный списки
                    {
                        case WdListType.wdListBullet:
                            paragraphInfo.Type = ParagraphInfo.ParagraphClass.ЭлементМаркерованногоСписка;
                            break;
                        case WdListType.wdListSimpleNumbering:
                            paragraphInfo.Type = ParagraphInfo.ParagraphClass.NumberList;
                            break;
                        case WdListType.wdListMixedNumbering:
                            paragraphInfo.Type = ParagraphInfo.ParagraphClass.NumberList;
                            break;

                        case WdListType.wdListOutlineNumbering:

                            string marker = paragraph.Range.ListFormat.ListString.Trim();
                            if (regexNumber.IsMatch(marker))
                            {
                                paragraphInfo.Type = ParagraphInfo.ParagraphClass.NumberList;
                            }
                            else
                            {
                                if (regexMarker.IsMatch(marker))
                                {
                                    paragraphInfo.Type = ParagraphInfo.ParagraphClass.ЭлементМаркерованногоСписка;
                                }
                                else
                                {   // неизвестный маркер
                                    paragraphInfo.Type = ParagraphInfo.ParagraphClass.ЭлементМаркерованногоСписка;
                                }
                            }

                            break;
                    }

                    infos.Add(paragraphInfo);
                    continue;
                }

                bool isCourier = false;
                // если текст из латинских букв - то это код
                if (!isCyrilic(paragraph.Range, out isCourier) || isCourier)
                {
                    paragraphInfo.Type = ParagraphInfo.ParagraphClass.Код;
                    infos.Add(paragraphInfo);
                    continue;
                }

                // анализ первого слова. так как это может быть заголовок с номером, нумерованный текст или заголовки объектов
                string firstWord = ExtractFirstWord(paragraph.Range);

                if (firstWord != null)
                {
                    if (regexNumber.IsMatch(firstWord[0] + "")) // если с цифры - то или заголовок или номер
                    {
                        paragraphInfo.Type = ParagraphInfo.ParagraphClass.NumberText;
                        infos.Add(paragraphInfo);
                        continue;
                    }

                    if (regexMarker.IsMatch(firstWord[0] + "")) // с маркера - маркированный текст
                    {
                        paragraphInfo.Type = ParagraphInfo.ParagraphClass.ЭлементМаркерованногоСписка;
                        infos.Add(paragraphInfo);
                        continue;
                    }

                    // TODO если абзац обычного текста начинается с фразы Листинг или Таблица или Рисунок - ложное срабатывание

                    if (firstWord.StartsWith("рисунок") || firstWord.StartsWith("рис.")) // подрисуночная подпись
                    {
                        paragraphInfo.Type = ParagraphInfo.ParagraphClass.ПодрисуночнаяПодпись;
                        infos.Add(paragraphInfo);
                        continue;
                    }

                    if (firstWord.StartsWith("листинг")) // название листинга
                    {
                        paragraphInfo.Type = ParagraphInfo.ParagraphClass.НазваниеЛистинга;
                        infos.Add(paragraphInfo);
                        continue;
                    }

                    if (firstWord.StartsWith("таблица") || firstWord.StartsWith("табл.")) // название таблицы
                    {
                        paragraphInfo.Type = ParagraphInfo.ParagraphClass.НазваниеТаблицы;
                        infos.Add(paragraphInfo);
                        continue;
                    }
                }

                if (text.StartsWith("список литературы") || // заголовок списка литературы
                    text.StartsWith("список источников") ||
                     text.StartsWith("список использованных источников") ||
                     text.StartsWith("список используемых источников"))
                {
                    paragraphInfo.Type = ParagraphInfo.ParagraphClass.ЗаголовокСпискаЛитературы;
                    documentParams.HasSource = true;
                    infos.Add(paragraphInfo);
                    continue;
                }

                paragraphInfo.Type = ParagraphInfo.ParagraphClass.Текст;

                infos.Add(paragraphInfo);
            }

            // копируем сюда список абзацев для обработки
            List<ParagraphInfo> buffer = new List<ParagraphInfo>(infos);

            int paragraphCount = document.Paragraphs.Count;
            for (int i = 1; i < infos.Count - 1; i++)
            {
                // пустые строки в коде помечаем как код
                if (infos[i - 1].Type == ParagraphInfo.ParagraphClass.Код &&
                    infos[i + 1].Type == ParagraphInfo.ParagraphClass.Код &&
                    infos[i].Type != ParagraphInfo.ParagraphClass.Код)
                {
                    if (infos[i].Type == ParagraphInfo.ParagraphClass.Empty ||
                        infos[i].Type == ParagraphInfo.ParagraphClass.Текст)
                    {
                        infos[i].Type = ParagraphInfo.ParagraphClass.Код;
                    }
                }
            }

            // удаляем пустые строки, чтобы не мешали анализу
            for (int i = buffer.Count - 1; i >= 0; i--)
            {
                if (buffer[i].Type == ParagraphInfo.ParagraphClass.Empty)
                {
                    buffer.RemoveAt(i);
                }
            }

            // если между двумя строками кода, есть кириллический текст или пустая строка - помечаем кодом
            for (int i = 1; i < buffer.Count - 1; i++)
            {
                if (buffer[i - 1].Type == ParagraphInfo.ParagraphClass.Код &&
                    buffer[i + 1].Type == ParagraphInfo.ParagraphClass.Код &&
                    buffer[i].Type != ParagraphInfo.ParagraphClass.Код)
                {
                    if (buffer[i].Type == ParagraphInfo.ParagraphClass.Empty || // TODO бессмысленное условие - глянь
                        buffer[i].Type == ParagraphInfo.ParagraphClass.Текст)
                    {
                        buffer[i].Type = ParagraphInfo.ParagraphClass.Код;
                    }
                }
            }

            // с учетом разметки простого текста обрабатываем заново пустые строки
            for (int i = 1; i < infos.Count - 1; i++)
            {
                if (infos[i - 1].Type == ParagraphInfo.ParagraphClass.Код &&
                    infos[i + 1].Type == ParagraphInfo.ParagraphClass.Код &&
                    infos[i].Type != ParagraphInfo.ParagraphClass.Код)
                {
                    if (infos[i].Type == ParagraphInfo.ParagraphClass.Empty ||
                        infos[i].Type == ParagraphInfo.ParagraphClass.Текст)
                    {
                        infos[i].Type = ParagraphInfo.ParagraphClass.Код;
                    }
                }
            }

            // считаем, что все пустые строки кода включили в код, тогда можно остальные удалять и пометить их

            for (int i = infos.Count - 1; i >= 0; i--)
            {
                if (infos[i].Type == ParagraphInfo.ParagraphClass.Empty)
                {
                    infos.RemoveAt(i);
                }
                else
                {
                    break; // TODO у последних пустых строк косячный Range, ставит примечания в произвольные места
                }
            }


            for (int i = 0; i < infos.Count; i++)
            {
                if (infos[i].Type == ParagraphInfo.ParagraphClass.Empty)
                {
                    var firstLiteralRangeComment = document.Paragraphs[infos[i].Index].Range;
                    firstLiteralRangeComment.End = firstLiteralRangeComment.Start + 1;

                    var noteComment = document.Comments.Add(firstLiteralRangeComment,
                        "# убрать пустые строки: отступы должны выполняться интервалами, а перенос на новый лист - свойством абзаца \"с новой страницы\"");
                    noteComment.Author = "ROBOT";
                }
            }

            for (int i = infos.Count - 1; i >= 0; i--)
            {
                if (infos[i].Type == ParagraphInfo.ParagraphClass.Empty)
                {
                    infos.RemoveAt(i);
                }
            }

            // в начале документа может быть две строки нумерованного текста заголовков - помечаем
            if (infos.Count > 2)
            {
                if (infos[0].Type == ParagraphInfo.ParagraphClass.NumberText || infos[0].Type == ParagraphInfo.ParagraphClass.NumberList)
                {
                    infos[0].Type = ParagraphInfo.ParagraphClass.Заголовок;
                    documentParams.HasTitle = true;
                }

                // TODO можно еще сравнить их уровни..

                if (infos[1].Type == ParagraphInfo.ParagraphClass.NumberText || infos[1].Type == ParagraphInfo.ParagraphClass.NumberList)
                {
                    infos[1].Type = ParagraphInfo.ParagraphClass.Заголовок;
                }
            }

            // далее считаем, что одиночный нумерованный элемент - заголовок, иначе - список 
            for (int i = 1; i < infos.Count - 1; i++)
            {
                bool topIsList = infos[i - 1].Type == ParagraphInfo.ParagraphClass.NumberList ||
                                  infos[i - 1].Type == ParagraphInfo.ParagraphClass.NumberText ||
                                  infos[i - 1].Type == ParagraphInfo.ParagraphClass.ЭлементНумерованногоСписка; // уже могли пометить

                bool currentIsList = infos[i].Type == ParagraphInfo.ParagraphClass.NumberList ||
                                     infos[i].Type == ParagraphInfo.ParagraphClass.NumberText ||
                                     infos[i].Type == ParagraphInfo.ParagraphClass.ЭлементНумерованногоСписка;

                bool bottomIsList = infos[i + 1].Type == ParagraphInfo.ParagraphClass.NumberList ||
                                    infos[i + 1].Type == ParagraphInfo.ParagraphClass.NumberText ||
                                    infos[i + 1].Type == ParagraphInfo.ParagraphClass.ЭлементНумерованногоСписка; // уже могли пометить

                if (currentIsList)
                {
                    if (!topIsList && !bottomIsList) // соседи не нумерованные - значит заголовок
                    {
                        infos[i].Type = ParagraphInfo.ParagraphClass.Заголовок;
                    }
                    else
                    { // иначе список - как и его друзья
                        infos[i].Type = ParagraphInfo.ParagraphClass.ЭлементНумерованногоСписка;

                        if (topIsList)
                        {
                            infos[i - 1].Type = ParagraphInfo.ParagraphClass.ЭлементНумерованногоСписка;
                        }

                        if (bottomIsList)
                        {
                            infos[i + 1].Type = ParagraphInfo.ParagraphClass.ЭлементНумерованногоСписка;
                        }

                        if (!bottomIsList)
                        {
                            infos[i].isLastListElement = true;
                        }

                        if (!topIsList)
                        {
                            if (infos[i - 1].Type == ParagraphInfo.ParagraphClass.Текст)
                            {
                                infos[i - 1].isTextBeforeList = true;
                            }
                            else
                            {
                                infos[i].Problems.Add("перед списком должен быть абзац текста, оканчивающийся двоеточием");
                            }
                        }
                    }
                }
            }

            for (int i = 1; i < infos.Count - 1; i++)
            {
                if (infos[i].Type == ParagraphInfo.ParagraphClass.ЭлементМаркерованногоСписка)
                {
                    if (infos[i - 1].Type == ParagraphInfo.ParagraphClass.Текст)
                    {
                        infos[i - 1].isTextBeforeList = true;
                    }
                    else
                    {
                        if (infos[i - 1].Type != ParagraphInfo.ParagraphClass.ЭлементМаркерованногоСписка)
                        {
                            infos[i].Problems.Add("перед списком должен быть абзац текста, оканчивающийся двоеточием");
                        }
                    }

                    if (infos[i + 1].Type != ParagraphInfo.ParagraphClass.ЭлементМаркерованногоСписка)
                    {
                        infos[i].isLastListElement = true;
                    }
                }
            }

            string levelNumber = null;
            int indexImage = 1;
            int indexCode = 1;
            int indexTable = 1;

            List<ParagraphInfo> references = new List<ParagraphInfo>();

            // проверяем наличие названий рисунков, листингов и таблиц
            for (int i = 1; i < infos.Count - 1; i++)
            {
                if (infos[i].Type == ParagraphInfo.ParagraphClass.Рисунок && infos[i + 1].Type != ParagraphInfo.ParagraphClass.ПодрисуночнаяПодпись)
                {
                    infos[i].Problems.Add("под рисунком должна быть подрисуночная подпись");
                }

                if (infos[i].Type == ParagraphInfo.ParagraphClass.Код
                    && infos[i + 1].Type != ParagraphInfo.ParagraphClass.НазваниеЛистинга
                    && infos[i + 1].Type != ParagraphInfo.ParagraphClass.Код)
                {
                    infos[i].Problems.Add("под листингом должна быть подпись (Листинг 1.1 – Название листинга)");
                    // TODO считаем, что толкьо под, хотя по требованиям АВ можно и над
                }

                if (infos[i].Type == ParagraphInfo.ParagraphClass.Таблица && infos[i - 1].Type != ParagraphInfo.ParagraphClass.НазваниеТаблицы
                    && infos[i - 1].Type != ParagraphInfo.ParagraphClass.Таблица)
                {
                    infos[i].Problems.Add("над таблицей быть название (подпись)");
                }
            }

            // обработка форматирования
            for (int i = 0; i < infos.Count; i++)
            {
                var paragraph = document.Paragraphs[infos[i].Index];
                application.Selection.SetRange(paragraph.Range.Start, paragraph.Range.End);

                switch (infos[i].Type)
                {
                    case ParagraphInfo.ParagraphClass.Текст:
                        {
                            if (paragraph.Format.Alignment != WdParagraphAlignment.wdAlignParagraphJustify)
                            {
                                infos[i].Problems.Add("установить выравнивание текста по ширине");
                            }

                            if (Math.Abs(paragraph.Format.FirstLineIndent - 35.45f) >= 0.1f)
                            {
                                infos[i].Problems.Add("добавить красную строку в 1,25 см");
                            }

                            if (paragraph.Format.LeftIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ слева");
                            }

                            if (paragraph.Format.RightIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ справа");
                            }

                            if (paragraph.Format.LineSpacingRule != WdLineSpacing.wdLineSpace1pt5)
                            {
                                infos[i].Problems.Add("установить полуторный межстрочный интервал");
                            }

                            if (paragraph.Format.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
                            {
                                infos[i].Problems.Add("убрать уровень текста");
                            }

                            // !((Microsoft.Office.Interop.Word.Style)paragraph.Range.get_Style()).NoSpaceBetweenParagraphsOfSameStyle
                            if (paragraph.Format.SpaceAfter != 0 || paragraph.Format.SpaceAfterAuto != 0)
                            {
                                infos[i].Problems.Add("убрать интервал после абзаца");
                            }

                            if ((paragraph.Format.SpaceBeforeAuto != 0 || paragraph.Format.SpaceBefore != 0) &&
                                (i != 0 && infos[i - 1].Type != ParagraphInfo.ParagraphClass.Таблица))
                            {
                                infos[i].Problems.Add("убрать интервал до абзаца");
                            }

                            bool[] problems = new bool[13];

                            foreach (Microsoft.Office.Interop.Word.Range word in paragraph.Range.Words)
                            {
                                if (word.Text.Trim().Length == 0)
                                {
                                    continue;
                                }

                                int indexProblem = 0;
                                bool hasProblem = false;

                                if (word.Font.Bold != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить нежирный шрифт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Italic != 0)
                                {
                                    var regex = new Regex("[^A-Za-z]");

                                    if (regex.IsMatch(word.Text))
                                    {
                                        if (!problems[indexProblem])
                                        {
                                            infos[i].Problems.Add("убрать курсив");
                                            problems[indexProblem] = true;
                                        }

                                        hasProblem = true;
                                    }

                                }

                                indexProblem++;

                                if (word.Font.StrikeThrough != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать зачеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Underline != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать подчеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.Subscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать подстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                //if (word.Font.Superscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать надстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                if (word.Font.ColorIndex != WdColorIndex.wdBlack && word.Font.ColorIndex != WdColorIndex.wdAuto)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить черный цвет шрифта");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.AllCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.SmallCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Fill)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Glow)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;
                                if (word.Font.Name == "")
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (word.Font.Name != "Times New Roman")
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить шрифт Times New Roman");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Size == 9999999)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (!(word.Font.Size == 14 || word.Font.Size == 12))
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить размер шрифта в 14 пт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (hasProblem)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdYellow;
                                }

                            }

                            checkText(paragraph, infos[i], documentParams);
                        }

                        break;
                    case ParagraphInfo.ParagraphClass.Код:
                        {
                            string text = paragraph.Range.Text.Replace('\a', ' ').Trim();

                            if (text.Length == 0)
                            {
                                continue;
                            }

                            if (paragraph.Format.Alignment != WdParagraphAlignment.wdAlignParagraphLeft)
                            {
                                infos[i].Problems.Add("установить выравнивание текста по левому краю");
                            }

                            if (paragraph.Format.FirstLineIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать красную строку");
                            }

                            //if (paragraph.Format.KeepTogether != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            //if (paragraph.Format.KeepWithNext != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            if (paragraph.Format.LeftIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ слева");
                            }

                            if (paragraph.Format.RightIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ справа");
                            }

                            if (paragraph.Format.LineSpacingRule != WdLineSpacing.wdLineSpaceSingle)
                            {
                                infos[i].Problems.Add("установить одинарный межстрочный интервал");
                            }

                            if (paragraph.Format.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
                            {
                                infos[i].Problems.Add("убрать уровень текста");
                            }

                            //if (paragraph.Format.PageBreakBefore != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            //if ((paragraph.Format.SpaceAfter != application.CentimetersToPoints(0) ||
                            //    paragraph.Format.SpaceBefore != application.CentimetersToPoints(0)) &&
                            //    !((Microsoft.Office.Interop.Word.Style)paragraph.Range.get_Style()).NoSpaceBetweenParagraphsOfSameStyle)
                            //{
                            //    infos[i].problems.Add("убрать интервал между абзацами");
                            //}

                            //if ((paragraph.Format.SpaceAfter != application.CentimetersToPoints(0) ||
                            //    paragraph.Format.SpaceBefore != application.CentimetersToPoints(0)) ||
                            //    paragraph.Format.SpaceAfterAuto == 1 || paragraph.Format.SpaceBeforeAuto == 1)
                            //{
                            //    infos[i].problems.Add("убрать интервал между абзацами");
                            //}                

                            bool[] problems = new bool[13];

                            var range = paragraph.Range;
                            if (paragraph.Range.Text.EndsWith("\a"))
                            {
                                range.End = range.End - 1; // символ \a имеет нестандартое форматирование шрифта и меняет определение шрифта всего абзаца
                            }
                            foreach (Microsoft.Office.Interop.Word.Range word in range.Words)
                            {
                                int indexProblem = 0;
                                bool hasProblem = false;

                                if (word.Font.Bold != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить нежирный шрифт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Italic != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать курсив");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.StrikeThrough != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать зачеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Underline != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать подчеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Subscript != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать подстрочный текст");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Superscript != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать надстрочный текст");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.ColorIndex != WdColorIndex.wdBlack && word.Font.ColorIndex != WdColorIndex.wdAuto)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить черный цвет шрифта");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.AllCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.SmallCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Fill)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Glow)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                if (word.Font.Name == "")
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (word.Font.Name != "Courier New")
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить допустимое семейство шрифтов (Courier New)");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Size == 9999999)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (!(word.Font.Size >= 9 && word.Font.Size <= 12))
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить меньший размер шрифта (9-12 пт)");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (hasProblem)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdYellow;
                                }
                            }

                        }

                        break;
                    case ParagraphInfo.ParagraphClass.Рисунок:
                        {
                            if (paragraph.Format.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
                            {
                                infos[i].Problems.Add("установить выравнивание текста по центру");
                            }

                            if (paragraph.Format.FirstLineIndent != 0)
                            {
                                infos[i].Problems.Add("убрать красную строку");
                            }

                            //if (paragraph.Format.KeepTogether != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            if (paragraph.Format.KeepWithNext != -1)
                            {
                                infos[i].Problems.Add("выставить свойство абзаца \'не отрывать от следующего\'");
                            }

                            if (paragraph.Format.LeftIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ слева");
                            }

                            if (paragraph.Format.RightIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ справа");
                            }

                            //if (paragraph.Format.LineSpacingRule != WdLineSpacing.wdLineSpace1pt5)
                            //{
                            //    infos[i].problems.Add("установить полуторный межстрочный интервал");
                            //}

                            if (paragraph.Format.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
                            {
                                infos[i].Problems.Add("убрать уровень текста");
                            }

                            //if (paragraph.Format.PageBreakBefore != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            // TODO добавить отступы


                            if (paragraph.Format.SpaceBefore != 6)
                            {
                                infos[i].Problems.Add("установить интервал до абзаца в 6 пт");
                            }
                            else
                            {
                                if (paragraph.Format.SpaceBeforeAuto == 1)
                                {
                                    infos[i].Problems.Add("установить интервал до абзаца в 6 пт");
                                }
                            }

                            if (paragraph.Format.SpaceAfter > 6 || paragraph.Format.SpaceAfterAuto == 1)
                            {
                                infos[i].Problems.Add("интервал после абзаца не должен превышать 6 пт");
                            }

                        }
                        break;
                    case ParagraphInfo.ParagraphClass.Таблица:
                        {
                            string text = paragraph.Range.Text.Replace('\a', ' ').Trim();

                            if (text.Length == 0)
                            {
                                continue;
                            }

                            if (!(paragraph.Format.Alignment == WdParagraphAlignment.wdAlignParagraphJustify || paragraph.Format.Alignment == WdParagraphAlignment.wdAlignParagraphLeft
                                || paragraph.Format.Alignment == WdParagraphAlignment.wdAlignParagraphCenter))
                            {
                                infos[i].Problems.Add("установить выравнивание текста по ширине или по левому краю");
                            }

                            if (paragraph.Format.FirstLineIndent != 0f)
                            {
                                infos[i].Problems.Add("убрать красную строку");
                            }

                            //if (paragraph.Format.KeepTogether != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            //if (paragraph.Format.KeepWithNext != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            if (paragraph.Format.LeftIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ слева");
                            }

                            if (paragraph.Format.RightIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ справа");
                            }

                            if (paragraph.Format.LineSpacingRule != WdLineSpacing.wdLineSpaceSingle)
                            {
                                infos[i].Problems.Add("установить одинарный межстрочный интервал");
                            }

                            if (paragraph.Format.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
                            {
                                infos[i].Problems.Add("убрать уровень текста");
                            }

                            //if (paragraph.Format.PageBreakBefore != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}


                            if ((paragraph.Format.SpaceAfter != 0 ||
                                 paragraph.Format.SpaceAfterAuto == 1))
                            {
                                infos[i].Problems.Add("убрать интервал после абзаца");
                            }

                            if (paragraph.Format.SpaceBefore != 0 || paragraph.Format.SpaceBeforeAuto == 1)
                            {
                                infos[i].Problems.Add("убрать интервал после абзаца");
                            }

                            bool[] problems = new bool[13];

                            var range = paragraph.Range;
                            if (paragraph.Range.Text.EndsWith("\a"))
                            {
                                range.End = range.End - 1; // символ \a имеет нестандартое форматирование шрифта и меняет определение шрифта всего абзаца
                            }
                            foreach (Microsoft.Office.Interop.Word.Range word in range.Words)
                            {
                                int indexProblem = 0;
                                bool hasProblem = false;

                                if (word.Font.Bold != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить нежирный шрифт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Italic != 0)
                                {
                                    var regex = new Regex("[^A-Za-z]");

                                    if (regex.IsMatch(word.Text))
                                    {
                                        if (!problems[indexProblem])
                                        {
                                            infos[i].Problems.Add("убрать курсив");
                                            problems[indexProblem] = true;
                                        }

                                        hasProblem = true;
                                    }

                                }

                                indexProblem++;

                                if (word.Font.StrikeThrough != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать зачеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Underline != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать подчеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.Subscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать подстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                //if (word.Font.Superscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать надстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                if (word.Font.ColorIndex != WdColorIndex.wdBlack && word.Font.ColorIndex != WdColorIndex.wdAuto)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить черный цвет шрифта");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.AllCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.SmallCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Fill)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Glow)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;
                                if (word.Font.Name == "")
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (word.Font.Name != "Times New Roman")
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить шрифт Times New Roman");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Size == 9999999)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (!(word.Font.Size == 14 || word.Font.Size == 12))
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить размер шрифта в 14 пт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (hasProblem)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdYellow;
                                }

                            }

                            //var range = paragraph.Range;
                            //var find = range.Find;
                            //find.ClearFormatting();
                            //find.MatchWildcards = true;
                            //find.Text = @"[[]([0-9]{1;3})[]]";

                            //while (find.Execute())
                            //{
                            //    if (range.End > paragraph.Range.End)
                            //    {
                            //        break;
                            //    }
                            //    range.Start -= 1;
                            //    range.End += 1;
                            //    range.HighlightColorIndex = WdColorIndex.wdGreen;                        
                            //    range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);                       
                            //}
                        }
                        break;
                    case ParagraphInfo.ParagraphClass.Заголовок:
                        {
                            if (paragraph.Format.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
                            {
                                infos[i].Problems.Add("установить выравнивание текста по центру");
                            }

                            if (paragraph.Format.FirstLineIndent != 0)
                            {
                                infos[i].Problems.Add("убрать красную строку");
                            }

                            //if (paragraph.Format.KeepTogether != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            if (paragraph.Format.KeepWithNext != -1)
                            {
                                infos[i].Problems.Add("установить свойство абзаца \"не отрывать от следующего\"");
                            }

                            if (paragraph.Format.LeftIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ слева");
                            }

                            if (paragraph.Format.RightIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ справа");
                            }

                            if (paragraph.Format.LineSpacingRule != WdLineSpacing.wdLineSpace1pt5)
                            {
                                infos[i].Problems.Add("установить полуторный межстрочный интервал");
                            }

                            if (!(paragraph.Format.OutlineLevel == WdOutlineLevel.wdOutlineLevel1 ||
                                paragraph.Format.OutlineLevel == WdOutlineLevel.wdOutlineLevel2 ||
                                paragraph.Format.OutlineLevel == WdOutlineLevel.wdOutlineLevel3))
                            {
                                infos[i].Problems.Add("установить уровень абзаца на Уровень 1-3 в зависимости от типа заголовка: раздел, подраздел, пункт");
                            }

                            //if (paragraph.Format.PageBreakBefore != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            if (!(paragraph.Format.SpaceAfter >= 12 &&
                                (paragraph.Format.SpaceAfter <= 14)))
                            {
                                infos[i].Problems.Add("установить интервал после абзаца в 12-14 пт");
                            }
                            else
                            {
                                if (paragraph.Format.SpaceAfterAuto == 1)
                                {
                                    infos[i].Problems.Add("установить интервал после абзаца в 12-14 пт");
                                }
                            }

                            if (!(paragraph.Format.SpaceBefore >= 12 &&
                                 (paragraph.Format.SpaceBefore <= 14)))
                            {
                                infos[i].Problems.Add("установить интервал до абзаца в 12-14 пт");
                            }
                            else
                            {

                                if (paragraph.Format.SpaceBeforeAuto == 1)
                                {
                                    infos[i].Problems.Add("установить интервал до абзаца в 12-14 пт");
                                }
                            }


                            bool[] problems = new bool[13];
                            int x = 0;
                            foreach (Microsoft.Office.Interop.Word.Range word in paragraph.Range.Words)
                            {
                                if (word.Text.Trim().Length == 0)
                                {
                                    continue;
                                }

                                int indexProblem = 0;
                                bool hasProblem = false;

                                if (word.Font.Bold == 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить жирный шрифт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Italic != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать курсив");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.StrikeThrough != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать зачеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Underline != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать подчеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.Subscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать подстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                //if (word.Font.Superscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать надстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                if (word.Font.ColorIndex != WdColorIndex.wdBlack && word.Font.ColorIndex != WdColorIndex.wdAuto)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить черный цвет шрифта");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.AllCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.SmallCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Fill)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Glow)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                if (word.Font.Name == "")
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (word.Font.Name != "Times New Roman")
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить шрифт Times New Roman");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Size == 9999999)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (!(word.Font.Size == 14 || word.Font.Size == 12))
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить размер шрифта в 14 пт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (hasProblem)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdYellow;
                                }
                            }

                            if (levelNumber == null)
                            {
                                levelNumber = checkHeader(paragraph, infos[i], documentParams);
                            }
                            else
                            {
                                checkHeader(paragraph, infos[i], documentParams);
                            }

                        }
                        break;
                    case ParagraphInfo.ParagraphClass.ПодрисуночнаяПодпись:
                        {
                            if (paragraph.Format.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
                            {
                                infos[i].Problems.Add("установить выравнивание текста по центру");
                            }

                            if (paragraph.Format.FirstLineIndent != 0)
                            {
                                infos[i].Problems.Add("убрать красную строку");
                            }

                            //if (paragraph.Format.KeepTogether != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            //if (paragraph.Format.KeepWithNext != 1)
                            //{
                            //    infos[i].problems.Add("установить свойство абзаца \"не отрывать от следующего\"");
                            //}

                            if (paragraph.Format.LeftIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ слева");
                            }

                            if (paragraph.Format.RightIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ справа");
                            }

                            if (paragraph.Format.LineSpacingRule != WdLineSpacing.wdLineSpaceSingle)
                            {
                                infos[i].Problems.Add("установить одинарный межстрочный интервал");
                            }

                            if (paragraph.Format.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
                            {
                                infos[i].Problems.Add("убрать уровень абзаца");
                            }

                            //if (paragraph.Format.PageBreakBefore != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            if (paragraph.Format.SpaceAfter != 18)
                            {
                                infos[i].Problems.Add("установить интервал после абзаца в 18 пт");
                            }
                            else
                            {
                                if (paragraph.Format.SpaceAfterAuto == 1)
                                {
                                    infos[i].Problems.Add("установить интервал после абзаца в 18 пт");
                                }
                            }

                            if (paragraph.Format.SpaceBefore != 0 || paragraph.Format.SpaceBeforeAuto == 1)
                            {
                                infos[i].Problems.Add("убрать интервал до абзаца");
                            }


                            bool[] problems = new bool[13];

                            foreach (Microsoft.Office.Interop.Word.Range word in paragraph.Range.Words)
                            {
                                if (word.Text.Trim().Length == 0)
                                {
                                    continue;
                                }

                                int indexProblem = 0;
                                bool hasProblem = false;

                                if (word.Font.Bold != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить нежирный шрифт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Italic != 0)
                                {
                                    var regex = new Regex("[^A-Za-z]");

                                    if (regex.IsMatch(word.Text))
                                    {
                                        if (!problems[indexProblem])
                                        {
                                            infos[i].Problems.Add("убрать курсив");
                                            problems[indexProblem] = true;
                                        }

                                        hasProblem = true;
                                    }
                                }

                                indexProblem++;

                                if (word.Font.StrikeThrough != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать зачеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Underline != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать подчеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.Subscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать подстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                //if (word.Font.Superscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать надстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                if (word.Font.ColorIndex != WdColorIndex.wdBlack && word.Font.ColorIndex != WdColorIndex.wdAuto)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить черный цвет шрифта");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.AllCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.SmallCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Fill)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Glow)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                if (word.Font.Name == "")
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (word.Font.Name != "Times New Roman")
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить шрифт Times New Roman");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Size == 9999999)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (!(word.Font.Size == 14 || word.Font.Size == 12))
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить размер шрифта в 12 пт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (hasProblem)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdYellow;
                                }
                            }

                            infos[i].NumberObjectInText = checkObjectTitle(
                                             paragraph,
                                             infos[i],
                                             documentParams,
                                             ObjectTitleStrings.FigureTitle,
                                             levelNumber,
                                             ref indexImage);

                            infos[i].IndexObject = indexImage;

                            if (infos[i].NumberObjectInText != null)
                            {

                                infos[i].HasRef = false;
                                references.Add(infos[i]);
                            }

                            indexImage++;
                        }
                        break;
                    case ParagraphInfo.ParagraphClass.НазваниеТаблицы:
                        {
                            if (paragraph.Format.Alignment != WdParagraphAlignment.wdAlignParagraphLeft)
                            {
                                infos[i].Problems.Add("установить выравнивание текста по левому краю");
                            }

                            if (paragraph.Format.FirstLineIndent != 0)
                            {
                                infos[i].Problems.Add("убрать красную строку");
                            }

                            //if (paragraph.Format.KeepTogether != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            if (paragraph.Format.KeepWithNext != -1)
                            {
                                infos[i].Problems.Add("установить свойство абзаца \"не отрывать от следующего\"");
                            }

                            if (paragraph.Format.LeftIndent != 0)
                            {
                                infos[i].Problems.Add("убрать отступ слева");
                            }

                            if (paragraph.Format.RightIndent != 0)
                            {
                                infos[i].Problems.Add("убрать отступ справа");
                            }

                            if (paragraph.Format.LineSpacingRule != WdLineSpacing.wdLineSpaceSingle)
                            {
                                infos[i].Problems.Add("установить одинарный межстрочный интервал");
                            }

                            if (paragraph.Format.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
                            {
                                infos[i].Problems.Add("убрать уровень абзаца");
                            }

                            //if (paragraph.Format.PageBreakBefore != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            if (paragraph.Format.SpaceBefore != 3)
                            {
                                infos[i].Problems.Add("установить интервал до абзаца в 3 пт");
                            }
                            else
                            {
                                if (paragraph.Format.SpaceBeforeAuto == 1)
                                {
                                    infos[i].Problems.Add("установить интервал до абзаца в 3 пт");
                                }
                            }
                            if (paragraph.Format.SpaceBefore != 3)
                            {
                                infos[i].Problems.Add("установить интервал после абзаца в 3 пт");
                            }
                            else
                            {
                                if (paragraph.Format.SpaceBeforeAuto == 1)
                                {
                                    infos[i].Problems.Add("установить интервал после абзаца в 3 пт");
                                }
                            }


                            bool[] problems = new bool[13];

                            foreach (Microsoft.Office.Interop.Word.Range word in paragraph.Range.Words)
                            {
                                if (word.Text.Trim().Length == 0)
                                {
                                    continue;
                                }

                                int indexProblem = 0;
                                bool hasProblem = false;

                                if (word.Font.Bold != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить нежирный шрифт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Italic != 0)
                                {
                                    var regex = new Regex("[^A-Za-z]");

                                    if (regex.IsMatch(word.Text))
                                    {
                                        if (!problems[indexProblem])
                                        {
                                            infos[i].Problems.Add("убрать курсив");
                                            problems[indexProblem] = true;
                                        }

                                        hasProblem = true;
                                    }
                                }

                                indexProblem++;

                                if (word.Font.StrikeThrough != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать зачеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Underline != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать подчеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.Subscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать подстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                //if (word.Font.Superscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать надстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                if (word.Font.ColorIndex != WdColorIndex.wdBlack && word.Font.ColorIndex != WdColorIndex.wdAuto)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить черный цвет шрифта");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.AllCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.SmallCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Fill)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Glow)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                if (word.Font.Name == "")
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (word.Font.Name != "Times New Roman")
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить шрифт Times New Roman");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Size == 9999999)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (!(word.Font.Size == 14 || word.Font.Size == 12))
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить размер шрифта в 12 пт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (hasProblem)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdYellow;
                                }
                            }

                            infos[i].NumberObjectInText = checkObjectTitle(
                                                        paragraph,
                                                        infos[i],
                                                         documentParams,
                                                         ObjectTitleStrings.TableTitle,
                                                        levelNumber,
                                                        ref indexTable);

                            infos[i].IndexObject = indexTable;

                            if (infos[i].NumberObjectInText != null)
                            {

                                infos[i].HasRef = false;
                                references.Add(infos[i]);
                            }

                            indexTable++;
                        }
                        break;
                    case ParagraphInfo.ParagraphClass.НазваниеЛистинга:
                        {
                            if (paragraph.Format.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
                            {
                                infos[i].Problems.Add("установить выравнивание текста по центру");
                            }

                            if (paragraph.Format.FirstLineIndent != 0)
                            {
                                infos[i].Problems.Add("убрать красную строку");
                            }

                            //if (paragraph.Format.KeepTogether != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            //if (paragraph.Format.KeepWithNext != 1)
                            //{
                            //    infos[i].problems.Add("установить свойство абзаца \"не отрывать от следующего\"");
                            //}

                            if (paragraph.Format.LeftIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ слева");
                            }

                            if (paragraph.Format.RightIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ справа");
                            }

                            if (paragraph.Format.LineSpacingRule != WdLineSpacing.wdLineSpaceSingle)
                            {
                                infos[i].Problems.Add("установить одинарный межстрочный интервал");
                            }

                            if (paragraph.Format.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
                            {
                                infos[i].Problems.Add("убрать уровень абзаца");
                            }

                            //if (paragraph.Format.PageBreakBefore != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            if (paragraph.Format.SpaceAfter != 18)
                            {
                                infos[i].Problems.Add("установить интервал после абзаца в 18 пт");
                            }
                            else
                            {
                                if (paragraph.Format.SpaceAfterAuto == 1)
                                {
                                    infos[i].Problems.Add("установить интервал после абзаца в 18 пт");
                                }
                            }

                            if (paragraph.Format.SpaceBefore != 0 || paragraph.Format.SpaceBeforeAuto == 1)
                            {
                                infos[i].Problems.Add("убрать интервал до абзаца");
                            }


                            bool[] problems = new bool[13];

                            foreach (Microsoft.Office.Interop.Word.Range word in paragraph.Range.Words)
                            {
                                if (word.Text.Trim().Length == 0)
                                {
                                    continue;
                                }

                                int indexProblem = 0;
                                bool hasProblem = false;

                                if (word.Font.Bold != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить нежирный шрифт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Italic != 0)
                                {
                                    var regex = new Regex("[^A-Za-z]");

                                    if (regex.IsMatch(word.Text))
                                    {
                                        if (!problems[indexProblem])
                                        {
                                            infos[i].Problems.Add("убрать курсив");
                                            problems[indexProblem] = true;
                                        }

                                        hasProblem = true;
                                    }
                                }

                                indexProblem++;

                                if (word.Font.StrikeThrough != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать зачеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Underline != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать подчеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.Subscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать подстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                //if (word.Font.Superscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать надстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                if (word.Font.ColorIndex != WdColorIndex.wdBlack && word.Font.ColorIndex != WdColorIndex.wdAuto)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить черный цвет шрифта");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.AllCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.SmallCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Fill)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Glow)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                if (word.Font.Name == "")
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (word.Font.Name != "Times New Roman")
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить шрифт Times New Roman");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Size == 9999999)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (!(word.Font.Size == 14 || word.Font.Size == 12))
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить размер шрифта в 12 пт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (hasProblem)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdYellow;
                                }
                            }

                            infos[i].NumberObjectInText = checkObjectTitle(
                                             paragraph,
                                             infos[i],
                                             documentParams,
                                             ObjectTitleStrings.CodeTitle,
                                             levelNumber,
                                             ref indexCode);

                            infos[i].IndexObject = indexCode;

                            if (infos[i].NumberObjectInText != null)
                            {
                                infos[i].HasRef = false;
                                references.Add(infos[i]);
                            }

                            indexCode++;
                        }
                        break;
                    case ParagraphInfo.ParagraphClass.ЭлементНумерованногоСписка:
                        {
                            if (paragraph.Format.Alignment != WdParagraphAlignment.wdAlignParagraphJustify)
                            {
                                infos[i].Problems.Add("установить выравнивание текста по ширине");
                            }

                            if (Math.Abs(paragraph.Format.FirstLineIndent - 35.45f) >= 0.1f)
                            {
                                infos[i].Problems.Add("добавить красную строку в 1,25 см");
                            }

                            //if (paragraph.Format.KeepTogether != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            //if (paragraph.Format.KeepWithNext != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            if (paragraph.Format.LeftIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ слева");
                            }

                            if (paragraph.Format.RightIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ справа");
                            }

                            if (paragraph.Format.LineSpacingRule != WdLineSpacing.wdLineSpace1pt5)
                            {
                                infos[i].Problems.Add("установить полуторный межстрочный интервал");
                            }

                            if (paragraph.Format.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
                            {
                                infos[i].Problems.Add("убрать уровень текста");
                            }

                            //if (paragraph.Format.PageBreakBefore != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            // !((Microsoft.Office.Interop.Word.Style)paragraph.Range.get_Style()).NoSpaceBetweenParagraphsOfSameStyle
                            if (((paragraph.Format.SpaceAfter != 0 ||
                                paragraph.Format.SpaceAfterAuto != 0)))
                            {
                                infos[i].Problems.Add("убрать интервал после абзаца");
                            }

                            if ((paragraph.Format.SpaceBeforeAuto != 0 ||
                                paragraph.Format.SpaceBefore != 0) && (i != 0 && infos[i - 1].Type != ParagraphInfo.ParagraphClass.Таблица))
                            {
                                infos[i].Problems.Add("убрать интервал до абзаца");
                            }

                            bool[] problems = new bool[13];

                            foreach (Microsoft.Office.Interop.Word.Range word in paragraph.Range.Words)
                            {
                                if (word.Text.Trim().Length == 0)
                                {
                                    continue;
                                }

                                int indexProblem = 0;
                                bool hasProblem = false;

                                if (word.Font.Bold != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить нежирный шрифт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Italic != 0)
                                {
                                    var regex = new Regex("[^A-Za-z]");

                                    if (regex.IsMatch(word.Text))
                                    {
                                        if (!problems[indexProblem])
                                        {
                                            infos[i].Problems.Add("убрать курсив");
                                            problems[indexProblem] = true;
                                        }

                                        hasProblem = true;
                                    }

                                }

                                indexProblem++;

                                if (word.Font.StrikeThrough != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать зачеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Underline != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать подчеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.Subscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать подстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                //if (word.Font.Superscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать надстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                if (word.Font.ColorIndex != WdColorIndex.wdBlack && word.Font.ColorIndex != WdColorIndex.wdAuto)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить черный цвет шрифта");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.AllCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.SmallCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Fill)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Glow)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                if (word.Font.Name == "")
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (word.Font.Name != "Times New Roman")
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить шрифт Times New Roman");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Size == 9999999)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (!(word.Font.Size == 14 || word.Font.Size == 12))
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить размер шрифта в 14 пт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (hasProblem)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdYellow;
                                }

                            }

                            checkText(paragraph, infos[i], documentParams);
                        }
                        break;
                    case ParagraphInfo.ParagraphClass.ЭлементМаркерованногоСписка:
                        {
                            if (paragraph.Format.Alignment != WdParagraphAlignment.wdAlignParagraphJustify)
                            {
                                infos[i].Problems.Add("установить выравнивание текста по ширине");
                            }

                            if (Math.Abs(paragraph.Format.FirstLineIndent - 35.45f) >= 0.1f)
                            {
                                infos[i].Problems.Add("добавить красную строку в 1,25 см");
                            }

                            //if (paragraph.Format.KeepTogether != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            //if (paragraph.Format.KeepWithNext != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            if (paragraph.Format.LeftIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ слева");
                            }

                            if (paragraph.Format.RightIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ справа");
                            }

                            if (paragraph.Format.LineSpacingRule != WdLineSpacing.wdLineSpace1pt5)
                            {
                                infos[i].Problems.Add("установить полуторный межстрочный интервал");
                            }

                            if (paragraph.Format.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
                            {
                                infos[i].Problems.Add("убрать уровень текста");
                            }

                            //if (paragraph.Format.PageBreakBefore != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            // !((Microsoft.Office.Interop.Word.Style)paragraph.Range.get_Style()).NoSpaceBetweenParagraphsOfSameStyle
                            if (((paragraph.Format.SpaceAfter != 0 ||
                                paragraph.Format.SpaceAfterAuto != 0)))
                            {
                                infos[i].Problems.Add("убрать интервал после абзаца");
                            }

                            if ((paragraph.Format.SpaceBeforeAuto != 0 ||
                                paragraph.Format.SpaceBefore != 0) && (i != 0 && infos[i - 1].Type != ParagraphInfo.ParagraphClass.Таблица))
                            {
                                infos[i].Problems.Add("убрать интервал до абзаца");
                            }

                            bool[] problems = new bool[13];

                            foreach (Microsoft.Office.Interop.Word.Range word in paragraph.Range.Words)
                            {
                                if (word.Text.Trim().Length == 0)
                                {
                                    continue;
                                }

                                int indexProblem = 0;
                                bool hasProblem = false;

                                if (word.Font.Bold != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить нежирный шрифт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Italic != 0)
                                {
                                    var regex = new Regex("[^A-Za-z]");

                                    if (regex.IsMatch(word.Text))
                                    {
                                        if (!problems[indexProblem])
                                        {
                                            infos[i].Problems.Add("убрать курсив");
                                            problems[indexProblem] = true;
                                        }

                                        hasProblem = true;
                                    }

                                }

                                indexProblem++;

                                if (word.Font.StrikeThrough != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать зачеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Underline != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать подчеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.Subscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать подстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                //if (word.Font.Superscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать надстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                if (word.Font.ColorIndex != WdColorIndex.wdBlack && word.Font.ColorIndex != WdColorIndex.wdAuto)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить черный цвет шрифта");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.AllCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.SmallCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Fill)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Glow)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                if (word.Font.Name != "Times New Roman")
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить шрифт Times New Roman");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Size == 9999999)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (!(word.Font.Size == 14 || word.Font.Size == 12))
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить размер шрифта в 14 пт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (hasProblem)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdYellow;
                                }

                            }

                            checkText(paragraph, infos[i], documentParams);
                        }
                        break;
                    case ParagraphInfo.ParagraphClass.БиблиографическоеОписаниеИсточника:
                        {
                            if (paragraph.Format.Alignment != WdParagraphAlignment.wdAlignParagraphJustify)
                            {
                                infos[i].Problems.Add("установить выравнивание текста по ширине");
                            }

                            if (Math.Abs(paragraph.Format.FirstLineIndent - 35.45f) >= 0.1f)
                            {
                                infos[i].Problems.Add("добавить красную строку в 1,25 см");
                            }

                            //if (paragraph.Format.KeepTogether != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            //if (paragraph.Format.KeepWithNext != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            if (paragraph.Format.LeftIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ слева");
                            }

                            if (paragraph.Format.RightIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ справа");
                            }

                            if (paragraph.Format.LineSpacingRule != WdLineSpacing.wdLineSpace1pt5)
                            {
                                infos[i].Problems.Add("установить полуторный межстрочный интервал");
                            }

                            if (paragraph.Format.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
                            {
                                infos[i].Problems.Add("убрать уровень текста");
                            }

                            //if (paragraph.Format.PageBreakBefore != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            // !((Microsoft.Office.Interop.Word.Style)paragraph.Range.get_Style()).NoSpaceBetweenParagraphsOfSameStyle
                            if (((paragraph.Format.SpaceAfter != 0 ||
                                paragraph.Format.SpaceAfterAuto != 0)))
                            {
                                infos[i].Problems.Add("убрать интервал после абзаца");
                            }

                            if ((paragraph.Format.SpaceBeforeAuto != 0 ||
                                paragraph.Format.SpaceBefore != 0) && (i != 0 && infos[i - 1].Type != ParagraphInfo.ParagraphClass.Таблица))
                            {
                                infos[i].Problems.Add("убрать интервал до абзаца");
                            }

                            bool[] problems = new bool[13];

                            foreach (Microsoft.Office.Interop.Word.Range word in paragraph.Range.Words)
                            {
                                if (word.Text.Trim().Length == 0)
                                {
                                    continue;
                                }

                                int indexProblem = 0;
                                bool hasProblem = false;

                                if (word.Font.Bold != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить нежирный шрифт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Italic != 0)
                                {
                                    var regex = new Regex("[^A-Za-z]");

                                    if (regex.IsMatch(word.Text))
                                    {
                                        if (!problems[indexProblem])
                                        {
                                            infos[i].Problems.Add("убрать курсив");
                                            problems[indexProblem] = true;
                                        }

                                        hasProblem = true;
                                    }

                                }

                                indexProblem++;

                                if (word.Font.StrikeThrough != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать зачеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Underline != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать подчеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.Subscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать подстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                //if (word.Font.Superscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать надстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                if (word.Font.ColorIndex != WdColorIndex.wdBlack && word.Font.ColorIndex != WdColorIndex.wdAuto)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить черный цвет шрифта");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.AllCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.SmallCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Fill)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Glow)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                if (word.Font.Name == "")
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (word.Font.Name != "Times New Roman")
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить шрифт Times New Roman");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Size == 9999999)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (!(word.Font.Size == 14 || word.Font.Size == 12))
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить размер шрифта в 14 пт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (hasProblem)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdYellow;
                                }
                            }

                            checkSource(paragraph, infos[i], documentParams);
                        }
                        break;
                    case ParagraphInfo.ParagraphClass.ЗаголовокСпискаЛитературы:
                        {
                            if (paragraph.Format.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
                            {
                                infos[i].Problems.Add("установить выравнивание текста по центру");
                            }

                            if (paragraph.Format.FirstLineIndent != 0)
                            {
                                infos[i].Problems.Add("убрать красную строку");
                            }

                            //if (paragraph.Format.KeepTogether != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            if (paragraph.Format.KeepWithNext != -1)
                            {
                                infos[i].Problems.Add("установить свойство абзаца \"не отрывать от следующего\"");
                            }

                            if (paragraph.Format.LeftIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ слева");
                            }

                            if (paragraph.Format.RightIndent != application.CentimetersToPoints(0))
                            {
                                infos[i].Problems.Add("убрать отступ справа");
                            }

                            if (paragraph.Format.LineSpacingRule != WdLineSpacing.wdLineSpace1pt5)
                            {
                                infos[i].Problems.Add("установить полуторный межстрочный интервал");
                            }

                            //if (!(paragraph.Format.OutlineLevel == WdOutlineLevel.wdOutlineLevel1 ||
                            //    paragraph.Format.OutlineLevel == WdOutlineLevel.wdOutlineLevel2 ||
                            //    paragraph.Format.OutlineLevel == WdOutlineLevel.wdOutlineLevel3))
                            //{
                            //    infos[i].problems.Add("установить уровень абзаца на Уровень 1-3 в зависимости от типа заголовка: раздел, подраздел, пункт");
                            //}

                            //if (paragraph.Format.PageBreakBefore != 1)
                            //{
                            //    infos[i].problems.Add("");
                            //}

                            if (!(paragraph.Format.SpaceAfter >= 12 &&
                                (paragraph.Format.SpaceAfter <= 14)))
                            {
                                infos[i].Problems.Add("установить интервал после абзаца в 12-14 пт");
                            }
                            else
                            {
                                if (paragraph.Format.SpaceAfterAuto == 1)
                                {
                                    infos[i].Problems.Add("установить интервал после абзаца в 12-14 пт");
                                }
                            }

                            if (!(paragraph.Format.SpaceBefore >= 12 &&
                                 (paragraph.Format.SpaceBefore <= 14)))
                            {
                                infos[i].Problems.Add("установить интервал до абзаца в 12-14 пт");
                            }
                            else
                            {

                                if (paragraph.Format.SpaceBeforeAuto == 1)
                                {
                                    infos[i].Problems.Add("установить интервал до абзаца в 12-14 пт");
                                }
                            }


                            bool[] problems = new bool[13];
                            int x = 0;
                            foreach (Microsoft.Office.Interop.Word.Range word in paragraph.Range.Words)
                            {
                                if (word.Text.Trim().Length == 0)
                                {
                                    continue;
                                }

                                int indexProblem = 0;
                                bool hasProblem = false;

                                if (word.Font.Bold == 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить жирный шрифт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Italic != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать курсив");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.StrikeThrough != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать зачеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Underline != 0)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("убрать подчеркивание");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.Subscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать подстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                //if (word.Font.Superscript != 0)
                                //{
                                //    if (!problems[indexProblem])
                                //    {
                                //        infos[i].problems.Add("убрать надстрочный текст");
                                //        problems[indexProblem] = true;
                                //    }

                                //    hasProblem = true;
                                //}

                                indexProblem++;

                                if (word.Font.ColorIndex != WdColorIndex.wdBlack && word.Font.ColorIndex != WdColorIndex.wdAuto)
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить черный цвет шрифта");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                //if (word.Font.AllCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.SmallCaps != 0)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Fill)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                //if (word.Font.Glow)
                                //{
                                //    infos[i].problems.Add("");
                                //}

                                indexProblem++;

                                if (word.Font.Name != "Times New Roman")
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить шрифт Times New Roman");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (word.Font.Size == 9999999)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdGray25;
                                }
                                else
                                if (!(word.Font.Size == 14 || word.Font.Size == 12))
                                {
                                    if (!problems[indexProblem])
                                    {
                                        infos[i].Problems.Add("установить размер шрифта в 14 пт");
                                        problems[indexProblem] = true;
                                    }

                                    hasProblem = true;
                                }

                                indexProblem++;

                                if (hasProblem)
                                {
                                    word.HighlightColorIndex = WdColorIndex.wdYellow;
                                }

                            }
                        }

                        break;
                }
            }

            // проверяем наличие ссылки в тексте на объекты
            for (int i = 0; i < infos.Count; i++)
            {
                if (infos[i].Type == ParagraphInfo.ParagraphClass.Текст
                    || infos[i].Type == ParagraphInfo.ParagraphClass.ЭлементНумерованногоСписка
                    || infos[i].Type == ParagraphInfo.ParagraphClass.ЭлементМаркерованногоСписка)
                {
                    string text = document.Paragraphs[infos[i].Index].Range.Text.Trim().ToLower();

                    for (int j = 0; j < references.Count; j++)
                    {
                        if (!references[j].HasRef && references[j].Index > infos[i].Index)
                        {
                            string objectname = "объект";
                            switch (references[j].Type)
                            {
                                case ParagraphInfo.ParagraphClass.ПодрисуночнаяПодпись:
                                    objectname = ObjectTitleStrings.FigureRef;
                                    break;
                                case ParagraphInfo.ParagraphClass.НазваниеТаблицы:
                                    objectname = ObjectTitleStrings.TableRef;
                                    break;
                                case ParagraphInfo.ParagraphClass.НазваниеЛистинга:
                                    objectname = ObjectTitleStrings.CodeRef;
                                    break;
                            }
                            var regex = new Regex(objectname + ".{2,20}" + references[j].NumberObjectInText);

                            if (regex.IsMatch(text))
                            {
                                references[j].HasRef = true;
                            }
                        }
                    }
                }
            }

            // указываем примечание, если на объект ссылка не обнаружена
            for (int i = 0; i < infos.Count; i++)
            {
                if (!infos[i].HasRef)
                {
                    string objectname = "объект";
                    switch (infos[i].Type)
                    {
                        case ParagraphInfo.ParagraphClass.ПодрисуночнаяПодпись:
                            objectname = ObjectTitleStrings.FigureTitle;
                            break;
                        case ParagraphInfo.ParagraphClass.НазваниеТаблицы:
                            objectname = ObjectTitleStrings.TableTitle;
                            break;
                        case ParagraphInfo.ParagraphClass.НазваниеЛистинга:
                            objectname = ObjectTitleStrings.CodeTitle;
                            break;
                    }

                    infos[i].Problems.Add($"отсутствует ссылка на {objectname} " + infos[i].NumberObjectInText + " перед объектом");
                }
            }
           
            string summaryComment = "# Общие замечания по документу:";

            // наличие списка литературы
            if (!documentParams.HasSource)
            {
                summaryComment += "\n- отсутствует блок списка литературы";
                documentParams.HasGeneralComments = true;
            }

            if (!documentParams.HasReference)
            {
                summaryComment += "\n- отсутствуют внутретекстовые ссылки на источники";
                documentParams.HasGeneralComments = true;
            }

            if (!documentParams.HasTitle)
            {
                summaryComment += "\n- отсутствует пронумерованный заголовок подраздела";
                documentParams.HasGeneralComments = true;
            }

            // поля документа
            if (Math.Abs(document.PageSetup.LeftMargin - (28.35f * 3f)) >= 0.1)
            {
                summaryComment += "\n- установить левое поле страницы в 30 мм";
                documentParams.HasGeneralComments = true;
            }

            if (Math.Abs(document.PageSetup.RightMargin - (28.35f * 1.5f)) >= 0.1)
            {
                summaryComment += "\n- установить правое поле страницы в 15 мм";
                documentParams.HasGeneralComments = true;
            }
            if (Math.Abs(document.PageSetup.TopMargin - (28.35f * 2f)) >= 0.1)
            {
                summaryComment += "\n- установить верхнее поле страницы в 20 мм";
                documentParams.HasGeneralComments = true;
            }

            if (Math.Abs(document.PageSetup.BottomMargin - (28.35f * 2f)) >= 0.1)
            {
                summaryComment += "\n- установить нижнее поле страницы в 20 мм";
                documentParams.HasGeneralComments = true;
            }

            if (document.Shapes.Count != 0)
            {
                summaryComment += "\n- рекомендуется вставлять картинки внутрь абзаца, сейчас они исключены из анализа";
                documentParams.HasGeneralComments = true;
            }

            // общие замечания по документу
            if (documentParams.HasGeneralComments)
            {
               var noteComment = document.Paragraphs[1].Range.Comments.Add(document.Paragraphs[1].Range.Words[1], summaryComment);
                noteComment.Author = "ROBOT";
            }

            // общие замечания по отдельным абзацам
            for (int i = 0; i < infos.Count; i++)
            {
                try
                {
                    if (infos[i].Problems.Count != 0)
                    {
                        string comment = "# " + infos[i].Type.ToString();

                        for (int j = 0; j < infos[i].Problems.Count; j++)
                        {
                            comment += "\r\n- " + infos[i].Problems[j];
                        }
                        var firstLiteralRangeComment = document.Paragraphs[infos[i].Index].Range;
                        firstLiteralRangeComment.End = firstLiteralRangeComment.Start + 1;
                        var noteComment = document.Comments.Add(firstLiteralRangeComment, comment);
                        noteComment.Author = "ROBOT";
                    }

                }
                catch { } // TODO есть скрытые пустые абзацы в таблицах - на них не ставит примечания, ирод
            }


            // Text             
            // диапазон чисел, числа от 10
            // отвязка рисунка
            // Sources
            // (отсутствие капслока) 
            // список второго уровня а),б),в)...
            // неразнывный пробел между номером и словом в ссылке на рисунок, листинг и таблицу


            document.Save();
            application.Quit();
        }


    }
}
