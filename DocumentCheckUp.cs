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
        public static class ObjectTitleMarker
        {
            public static String FigTitle { get { return "Рисунок"; } }
            public static String TableTitle { get { return "Таблица"; } }
            public static String CodeTitle { get { return "Листинг"; } }

            public static String FigRef { get { return "рис"; } }
            public static String TableRef { get { return "табл"; } }
            public static String CodeRef { get { return "лист"; } }
        }

        private class ParagraphInfo
        {
            public int Index = 0;
            public enum ParagraphType
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

            public ParagraphType Type = ParagraphType.Empty;

            public List<string> Problems = new List<string>();


            public int IndexObject = 0;
            public string NumberObjectInText = null;
            public bool HasRef = true;

            public bool isLastListElement = false;
            public bool isTextBeforeList = false;

            public bool isCellTableA = false;
        }

        private class DocumentParams
        {
            public bool HasTitle = false;
            public bool HasSource = false;
            public bool HasReference = false;

            public bool HasGeneralComments = false;
        }
        private static bool isCyrilic(Microsoft.Office.Interop.Word.Range range)
        {
            var regex = new System.Text.RegularExpressions.Regex("[а-я]");
            int cyrrilicCount = 0;
            int alphaWordCount = 0;
            var alpha = new Regex("[a-zа-я]");

            for (int i = 1; i <= range.Words.Count; i++)
            {
                string word = range.Words[i].Text.Trim().ToLower();

                if (alpha.IsMatch(word)) // считаем,сколько букв
                {
                    int n = regex.Matches(word).Count;
                    if (n > 0.6 * word.Length) // если более 60% - кириллическое слово
                    {
                        cyrrilicCount++;
                    }

                    alphaWordCount++;
                }
            }

            return (100.0 * cyrrilicCount / alphaWordCount > 60);
        }

        private static string GetFirstWord(Microsoft.Office.Interop.Word.Range range)
        {
            var regex = new System.Text.RegularExpressions.Regex("[1-9a-zа-я\\u2022\\u25aa\\u006f\\u2014\\u2013\\u202d]");
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
        private static void checkSource(Paragraph paragraph, ParagraphInfo paragraphInfo, DocumentParams documentParams)
        {
            var paragraphRange = paragraph.Range;
            string text = paragraphRange.Text;
            text = text.Trim();


            if (text.Length != 0 && text[text.Length - 1] != '.')
            {
                paragraphInfo.Problems.Add("должна быть точка в конце абзаца");
            }

            string[] denyURLs = new string[] { "wikipedia.org", "habr.com"  };
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

        private static void checkText(Paragraph paragraph, ParagraphInfo paragraphInfo, DocumentParams documentParams)
        {

            string text = paragraph.Range.Text;

            bool isV1 = false;
            bool isV2 = false;
            bool isV3 = false;

            if (paragraphInfo.Type == ParagraphInfo.ParagraphType.ЭлементНумерованногоСписка)
            {
                if (paragraph.Range.ListParagraphs.Count != 0)
                {
                    string marker = paragraph.Range.ListFormat.ListString.Trim();

                    text = text.Trim();

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

            if (paragraphInfo.Type == ParagraphInfo.ParagraphType.ЭлементМаркерованногоСписка)
            {
                isV2 = true;

                if (paragraph.Range.ListParagraphs.Count == 0)
                {
                    text = text.Trim();

                    var regexMarker = new Regex("[\\u2022\\u25aa\\u006f\\u2014\\u2013\\u202d]");

                    var regexMarkerWithWS = new Regex("[\\u2022\\u25aa\\u006f\\u2014\\u2013\\u202d][\\s]");

                    var regexMarkerAdvance = new Regex("[\u2013\u2022\u25aa]");

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

            if (paragraphInfo.Type == ParagraphInfo.ParagraphType.Текст)
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
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }

            regex = new Regex("([^a-zA-Zа-яА-Я])(-|—)[^a-zA-Zа-яА-Я>]");

            if (regex.IsMatch(text))
            {
                paragraphInfo.Problems.Add("использовать правильно тире – вместо дефиса - и длинного тире —");

                var matches = regex.Matches(text);

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
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }


            regex = new Regex("\\[[0-9]{1,3}\\]"); // TODO изменить регулярку с учетом 01 или 001

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
                            range.HighlightColorIndex = WdColorIndex.wdYellow;
                            range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                        }
                    }
                }
                else
                {   // используемый пробел не неразрывный!
                    if (!variants[2].IsMatch(text))
                    {
                        var matches = variants[3].Matches(text);

                        paragraphInfo.Problems.Add("между ссылкой на источник и словом должен быть неразрывный пробел (shift+ctrl+пробел)");

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
                                range.HighlightColorIndex = WdColorIndex.wdYellow;
                                range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                            }
                        }
                    }
                }


                if (variants[1].IsMatch(text))
                {
                    var matches = variants[1].Matches(text);

                    paragraphInfo.Problems.Add("ссылка на источник входит в предложение, поэтому точка ставится после ссылки");

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
                            range.HighlightColorIndex = WdColorIndex.wdYellow;
                            range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                        }
                    }
                }
            }

            regex = new Regex("[\"“‟”„''‚]");
            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);

                paragraphInfo.Problems.Add("необходимо использовать «кавычки-ёлочки», а не кавычки-лапки и одиночные кавычки");

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
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }

            regex = new Regex("[^«][a-zA-Zа-яА-Я]*_[a-zA-Zа-яА-Я]*[^»]");

            if (regex.IsMatch(text))
            {
                // нашли слово, которое следует обернуть в кавычки
                var matches = regex.Matches(text);

                paragraphInfo.Problems.Add("латинские названия переменных следует обернуть в кавычки-елочки");

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
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }

            regex = new Regex("[^«][a-zA-Zа-яА-Я]+::[a-zA-Zа-яА-Я]+[^»]");

            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);

                paragraphInfo.Problems.Add("латинские названия с пространством имен следует обернуть в кавычки-елочки");

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
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }

            regex = new Regex("[^«]([+*&^$~><=]|([\\+]{2})|([-]{2})|(->)|(>>)|(<<)|([\\+]=)|(-=)|([\\*]=)|(\\/=)|(>=)|(<=)|(!=)|(&&)|(::)|([\\|]{2}))[^»]");

            if (regex.IsMatch(text))
            {
                bool isNotCpp = false;
                var matches = regex.Matches(text);

                for (int i = 0; i < matches.Count; i++)
                {
                    if (matches[i].Value.ToLower().StartsWith("c+") || matches[i].Value.ToLower().StartsWith("с+"))
                    {
                        continue;
                    }

                    isNotCpp = true;

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

                if (isNotCpp)
                {
                    paragraphInfo.Problems.Add("знаки операторов следует обернуть в кавычки-елочки");
                }
            }

            regex = new Regex("[A-ZА-Я]{2,}");

            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);

                paragraphInfo.Problems.Add("напоминание: не забывайте расшифровать аббревиатуру в месте ее первого использования, например, программное обеспечение (ПО)");

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
                        range.HighlightColorIndex = WdColorIndex.wdGray25;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }

            regex = new Regex("([A-ZА-Я]+[^a-zA-Zа-яА-Я]){3,}");

            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);

                paragraphInfo.Problems.Add("не должно быть капсола в тексте (или у Вас три подряд аббревиатуры - так можно)");

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
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }

            regex = new Regex(".[\\s]{2,}.");

            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);

                paragraphInfo.Problems.Add("обнаружен множественный пробел, нужно сократить до одного");

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
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }

            bool denyHyperlinks = false;

            if (paragraph.Range.Hyperlinks.Count != 0)
            {
                for (int i = 1; i <= paragraph.Range.Hyperlinks.Count; i++)
                {
                    string url = paragraph.Range.Hyperlinks[i].Address;

                    if (url != null)
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

                        var newRange = paragraph.Range;
                        newRange.Start = range.Start - 1;
                        newRange.End = range.End +1;

                        newRange.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }
        }

        static string checkHeader(Paragraph paragraph, ParagraphInfo paragraphInfo, DocumentParams documentParams)
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
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }

            regex = new Regex("([^a-zA-Zа-яА-Я])(-|—)[^a-zA-Zа-яА-Я]");

            if (regex.IsMatch(text))
            {
                paragraphInfo.Problems.Add("использовать правильно тире – вместо дефиса - и длинного тире —");

                var matches = regex.Matches(text);

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
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }

            regex = new Regex("[\"“‟”„''‚]");
            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);

                paragraphInfo.Problems.Add("необходимо использовать «кавычки-ёлочки», а не кавычки-лапки и одиночные кавычки");

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
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
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


        static string checkObjectTitle(Paragraph paragraph, ParagraphInfo paragraphInfo, DocumentParams documentParams, string marker, string level, int number)
        {
            string text = paragraph.Range.Text.Trim();

            var regex = new Regex(".[\\s]{2,}.");

            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);

                paragraphInfo.Problems.Add("обнаружен множественный пробел, нужно сократить до одного");

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
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }

            regex = new Regex("([^a-zA-Zа-яА-Я])(-|—)[^a-zA-Zа-яА-Я]");

            if (regex.IsMatch(text))
            {
                paragraphInfo.Problems.Add("использовать правильно тире – вместо дефиса - и длинного тире —");

                var matches = regex.Matches(text);

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
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }

            regex = new Regex("[\"“‟”„''‚]");
            if (regex.IsMatch(text))
            {
                var matches = regex.Matches(text);

                paragraphInfo.Problems.Add("необходимо использовать «кавычки-ёлочки», а не кавычки-лапки и одиночные кавычки");

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
                        range.HighlightColorIndex = WdColorIndex.wdYellow;
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
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
                        paragraphInfo.Problems.Add($"не совпадает номер объекта с ожидаемым: {marker} {level}.{number}");

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
            DocumentParams documentParams = new DocumentParams();

            var application = new Microsoft.Office.Interop.Word.Application();
            application.Visible = true;

            Document document = application.Documents.Open(pathDoc, false);

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

                    paragraphInfo.Type = ParagraphInfo.ParagraphType.Рисунок;
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
                    paragraphInfo.Type = ParagraphInfo.ParagraphType.БиблиографическоеОписаниеИсточника;
                    infos.Add(paragraphInfo);
                    continue;
                }

                // если есть таблица - или таблица или код в рамке (в старых работах)
                if (paragraph.Range.Tables.Count != 0)
                {
                    string textTable = paragraph.Range.Text.Replace("\a", "").Trim();

                    if (textTable.Length == 0)
                    {
                        paragraphInfo.Type = ParagraphInfo.ParagraphType.Таблица;
                        paragraphInfo.isCellTableA = true;
                        infos.Add(paragraphInfo);
                        continue;
                    }
                    // проверить длину текста
                    if (!isCyrilic(paragraph.Range))
                    {
                        paragraphInfo.Type = ParagraphInfo.ParagraphType.Код;
                        infos.Add(paragraphInfo);
                        continue;
                    }

                    paragraphInfo.Type = ParagraphInfo.ParagraphType.Таблица;
                    infos.Add(paragraphInfo);
                    continue;
                }



                // регулярки на определение маркированного или нумерованного текста
                var regexNumber = new System.Text.RegularExpressions.Regex("[1-9]");
                var regexMarker = new System.Text.RegularExpressions.Regex("[\\u2022\\u25aa\\u006f\\u2014\\u2013\\u202d]");

                if (paragraph.Range.ListParagraphs.Count != 0) // если есть список
                {
                    switch (paragraph.Range.ListFormat.ListType) // разделение на маркерованный и нумерованный списки
                    {
                        case WdListType.wdListBullet:
                            paragraphInfo.Type = ParagraphInfo.ParagraphType.ЭлементМаркерованногоСписка;
                            break;
                        case WdListType.wdListSimpleNumbering:
                            paragraphInfo.Type = ParagraphInfo.ParagraphType.NumberList;
                            break;
                        case WdListType.wdListMixedNumbering:
                            paragraphInfo.Type = ParagraphInfo.ParagraphType.NumberList;
                            break;

                        case WdListType.wdListOutlineNumbering:

                            string marker = paragraph.Range.ListFormat.ListString.Trim();
                            if (regexNumber.IsMatch(marker))
                            {
                                paragraphInfo.Type = ParagraphInfo.ParagraphType.NumberList;
                            }
                            else
                            {
                                if (regexMarker.IsMatch(marker))
                                {
                                    paragraphInfo.Type = ParagraphInfo.ParagraphType.ЭлементМаркерованногоСписка;
                                }
                                else
                                {   // неизвестный маркер
                                    paragraphInfo.Type = ParagraphInfo.ParagraphType.ЭлементМаркерованногоСписка;
                                }
                            }

                            break;
                    }

                    infos.Add(paragraphInfo);
                    continue;
                }

                // если текст из латинских букв - то это код
                if (!isCyrilic(paragraph.Range))
                {
                    paragraphInfo.Type = ParagraphInfo.ParagraphType.Код;
                    infos.Add(paragraphInfo);
                    continue;
                }

                // анализ первого слова. так как это может быть заголовок с номером, нумерованный текст или заголовки объектов
                string firstWord = GetFirstWord(paragraph.Range);

                if (firstWord != null)
                {
                    if (regexNumber.IsMatch(firstWord[0] + "")) // если с цифры - то или заголовок или номер
                    {
                        paragraphInfo.Type = ParagraphInfo.ParagraphType.NumberText;
                        infos.Add(paragraphInfo);
                        continue;
                    }

                    if (regexMarker.IsMatch(firstWord[0] + "")) // с маркера - маркированный текст
                    {
                        paragraphInfo.Type = ParagraphInfo.ParagraphType.ЭлементМаркерованногоСписка;
                        infos.Add(paragraphInfo);
                        continue;
                    }

                    // TODO если абзац обычного текста начинается с фразы Листинг или Таблица или Рисунок - ложное срабатывание

                    if (firstWord.StartsWith("рисунок") || firstWord.StartsWith("рис.")) // подрисуночная подпись
                    {
                        paragraphInfo.Type = ParagraphInfo.ParagraphType.ПодрисуночнаяПодпись;
                        infos.Add(paragraphInfo);
                        continue;
                    }

                    if (firstWord.StartsWith("листинг")) // название листинга
                    {
                        paragraphInfo.Type = ParagraphInfo.ParagraphType.НазваниеЛистинга;
                        infos.Add(paragraphInfo);
                        continue;
                    }

                    if (firstWord.StartsWith("таблица") || firstWord.StartsWith("табл.")) // название таблицы
                    {
                        paragraphInfo.Type = ParagraphInfo.ParagraphType.НазваниеТаблицы;
                        infos.Add(paragraphInfo);
                        continue;
                    }
                }

                if (text.StartsWith("список литературы") || // заголовок списка литературы
                    text.StartsWith("список источников") ||
                     text.StartsWith("список использованных источников") ||
                     text.StartsWith("список используемых источников"))
                {
                    paragraphInfo.Type = ParagraphInfo.ParagraphType.ЗаголовокСпискаЛитературы;
                    documentParams.HasSource = true;
                    infos.Add(paragraphInfo);
                    continue;
                }

                paragraphInfo.Type = ParagraphInfo.ParagraphType.Текст;

                infos.Add(paragraphInfo);
            }

            // копируем сюда список абзацев для обработки
            List<ParagraphInfo> buffer = new List<ParagraphInfo>(infos);

            int paragraphCount = document.Paragraphs.Count;
            for (int i = 1; i < infos.Count - 1; i++)
            {
                // пустые строки в коде помечаем как код
                if (infos[i - 1].Type == ParagraphInfo.ParagraphType.Код &&
                    infos[i + 1].Type == ParagraphInfo.ParagraphType.Код &&
                    infos[i].Type != ParagraphInfo.ParagraphType.Код)
                {
                    if (infos[i].Type == ParagraphInfo.ParagraphType.Empty ||
                        infos[i].Type == ParagraphInfo.ParagraphType.Текст)
                    {
                        infos[i].Type = ParagraphInfo.ParagraphType.Код;
                    }
                }               
            }

            // удаляем пустые строки, чтобы не мешали анализу
            for (int i = buffer.Count - 1; i >= 0; i--)
            {
                if (buffer[i].Type == ParagraphInfo.ParagraphType.Empty)
                {
                    buffer.RemoveAt(i);
                }
            }

            // если между двумя строками кода, есть кириллический текст или пустая строка - помечаем кодом
            for (int i = 1; i < buffer.Count - 1; i++)
            {
                if (buffer[i - 1].Type == ParagraphInfo.ParagraphType.Код &&
                    buffer[i + 1].Type == ParagraphInfo.ParagraphType.Код &&
                    buffer[i].Type != ParagraphInfo.ParagraphType.Код)
                {
                    if (buffer[i].Type == ParagraphInfo.ParagraphType.Empty || // TODO бессмысленное условие - глянь
                        buffer[i].Type == ParagraphInfo.ParagraphType.Текст)
                    {
                        buffer[i].Type = ParagraphInfo.ParagraphType.Код;
                    }
                }
            }

            // с учетом разметки простого текста обрабатываем заново пустые строки
            for (int i = 1; i < infos.Count - 1; i++)
            {
                if (infos[i - 1].Type == ParagraphInfo.ParagraphType.Код &&
                    infos[i + 1].Type == ParagraphInfo.ParagraphType.Код &&
                    infos[i].Type != ParagraphInfo.ParagraphType.Код)
                {
                    if (infos[i].Type == ParagraphInfo.ParagraphType.Empty ||
                        infos[i].Type == ParagraphInfo.ParagraphType.Текст)
                    {
                        infos[i].Type = ParagraphInfo.ParagraphType.Код;
                    }
                }
            }

            // считаем, что все пустые строки кода включили в код, тогда можно остальные удалять и пометить их

            for (int i = infos.Count - 1; i >= 0; i--)
            {
                if (infos[i].Type == ParagraphInfo.ParagraphType.Empty)
                {
                    infos.RemoveAt(i);
                }
                else
                {
                    break; //TODO у последних пустых строк косячный Range, ставит примечания в произвольные места
                }
            }


            for (int i = 0; i < infos.Count ; i++)
            {
                if (infos[i].Type == ParagraphInfo.ParagraphType.Empty)
                {
                    var firstLiteralRangeComment = document.Paragraphs[infos[i].Index].Range;
                    firstLiteralRangeComment.End = firstLiteralRangeComment.Start + 1;

                    document.Comments.Add(firstLiteralRangeComment,
                        "убрать пустые строки: отступы должны выполняться интервалами, а перенос на новый лист - свойством абзаца \"с новой страницы\"");
                }
            }

            for (int i = infos.Count - 1; i >= 0; i--)
            {
                if (infos[i].Type == ParagraphInfo.ParagraphType.Empty)
                {
                    infos.RemoveAt(i);
                }
            }

            // в начале документа может быть две строки нумерованного текста заголовков - помечаем
            if (infos.Count > 2)
            {
                if (infos[0].Type == ParagraphInfo.ParagraphType.NumberText || infos[0].Type == ParagraphInfo.ParagraphType.NumberList)
                {
                    infos[0].Type = ParagraphInfo.ParagraphType.Заголовок;
                    documentParams.HasTitle = true;
                }

                // TODO можно еще сравнить их уровни..

                if (infos[1].Type == ParagraphInfo.ParagraphType.NumberText || infos[1].Type == ParagraphInfo.ParagraphType.NumberList)
                {
                    infos[1].Type = ParagraphInfo.ParagraphType.Заголовок;
                }
            }

            // далее считаем, что одиночный нумерованный элемент - заголовок, иначе - список 
            for (int i = 1; i < infos.Count - 1; i++)
            {
                bool topIsList = infos[i - 1].Type == ParagraphInfo.ParagraphType.NumberList ||
                                  infos[i - 1].Type == ParagraphInfo.ParagraphType.NumberText ||
                                  infos[i - 1].Type == ParagraphInfo.ParagraphType.ЭлементНумерованногоСписка; // уже могли пометить

                bool currentIsList = infos[i].Type == ParagraphInfo.ParagraphType.NumberList ||
                                     infos[i].Type == ParagraphInfo.ParagraphType.NumberText ||
                                     infos[i].Type == ParagraphInfo.ParagraphType.ЭлементНумерованногоСписка;

                bool bottomIsList = infos[i + 1].Type == ParagraphInfo.ParagraphType.NumberList ||
                                    infos[i + 1].Type == ParagraphInfo.ParagraphType.NumberText ||
                                    infos[i + 1].Type == ParagraphInfo.ParagraphType.ЭлементНумерованногоСписка; // уже могли пометить

                if (currentIsList)
                {
                    if (!topIsList && !bottomIsList) // соседи не нумерованные - значит заголовок
                    {
                        infos[i].Type = ParagraphInfo.ParagraphType.Заголовок;
                    }
                    else
                    { // иначе список - как и его друзья
                        infos[i].Type = ParagraphInfo.ParagraphType.ЭлементНумерованногоСписка;

                        if (topIsList)
                        {
                            infos[i - 1].Type = ParagraphInfo.ParagraphType.ЭлементНумерованногоСписка;
                        }

                        if (bottomIsList)
                        {
                            infos[i + 1].Type = ParagraphInfo.ParagraphType.ЭлементНумерованногоСписка;
                        }

                        if (!bottomIsList)
                        {
                            infos[i].isLastListElement = true;
                        }

                        if (!topIsList)
                        {
                            if (infos[i - 1].Type == ParagraphInfo.ParagraphType.Текст)
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
                if (infos[i].Type == ParagraphInfo.ParagraphType.ЭлементМаркерованногоСписка)
                {
                    if (infos[i - 1].Type == ParagraphInfo.ParagraphType.Текст)
                    {
                        infos[i - 1].isTextBeforeList = true;
                    }
                    else
                    {
                        if (infos[i - 1].Type != ParagraphInfo.ParagraphType.ЭлементМаркерованногоСписка)
                        {
                            infos[i].Problems.Add("перед списком должен быть абзац текста, оканчивающийся двоеточием");
                        }
                    }

                    if (infos[i + 1].Type != ParagraphInfo.ParagraphType.ЭлементМаркерованногоСписка)
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
                if (infos[i].Type == ParagraphInfo.ParagraphType.Рисунок && infos[i + 1].Type != ParagraphInfo.ParagraphType.ПодрисуночнаяПодпись)
                {
                    infos[i].Problems.Add("под рисунком должна быть подрисуночная подпись");
                }

                if (infos[i].Type == ParagraphInfo.ParagraphType.Код
                    && infos[i + 1].Type != ParagraphInfo.ParagraphType.НазваниеЛистинга
                    && infos[i + 1].Type != ParagraphInfo.ParagraphType.Код)
                {
                    infos[i].Problems.Add("под листингом должна быть подпись (Листинг 1.1 – Название листинга)");
                    // TODO считаем, что толкьо под, хотя по требованиям АВ можно и над
                }

                if (infos[i].Type == ParagraphInfo.ParagraphType.Таблица && infos[i - 1].Type != ParagraphInfo.ParagraphType.НазваниеТаблицы
                    && infos[i - 1].Type != ParagraphInfo.ParagraphType.Таблица)
                {
                    infos[i].Problems.Add("над таблицей быть название (подпись)");
                }
            }

            for (int i = 0; i < infos.Count; i++)
            {  
                var paragraph = document.Paragraphs[infos[i].Index];
                application.Selection.SetRange(paragraph.Range.Start, paragraph.Range.End);

                switch (infos[i].Type)
                {
                    case ParagraphInfo.ParagraphType.Текст:
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
                                paragraph.Format.SpaceBefore != 0) &&
                                (i != 0 && infos[i - 1].Type != ParagraphInfo.ParagraphType.Таблица))
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
                    case ParagraphInfo.ParagraphType.Код:
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
                    case ParagraphInfo.ParagraphType.Рисунок:
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
                    case ParagraphInfo.ParagraphType.Таблица:
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
                    case ParagraphInfo.ParagraphType.Заголовок:
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
                    case ParagraphInfo.ParagraphType.ПодрисуночнаяПодпись:
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
                                             ObjectTitleMarker.FigTitle,
                                             levelNumber,
                                             indexImage);

                            infos[i].IndexObject = indexImage;

                            if (infos[i].NumberObjectInText != null)
                            {

                                infos[i].HasRef = false;
                                references.Add(infos[i]);
                            }

                            indexImage++;
                        }
                        break;
                    case ParagraphInfo.ParagraphType.НазваниеТаблицы:
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
                                                         ObjectTitleMarker.TableTitle,
                                                        levelNumber,
                                                        indexTable);

                            infos[i].IndexObject = indexTable;

                            if (infos[i].NumberObjectInText != null)
                            {

                                infos[i].HasRef = false;
                                references.Add(infos[i]);
                            }

                            indexTable++;
                        }
                        break;
                    case ParagraphInfo.ParagraphType.НазваниеЛистинга:
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
                                             ObjectTitleMarker.CodeTitle,
                                             levelNumber,
                                             indexCode);

                            infos[i].IndexObject = indexCode;

                            if (infos[i].NumberObjectInText != null)
                            {
                                infos[i].HasRef = false;
                                references.Add(infos[i]);
                            }

                            indexCode++;
                        }
                        break;
                    case ParagraphInfo.ParagraphType.ЭлементНумерованногоСписка:
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
                                paragraph.Format.SpaceBefore != 0) && (i != 0 && infos[i - 1].Type != ParagraphInfo.ParagraphType.Таблица))
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
                    case ParagraphInfo.ParagraphType.ЭлементМаркерованногоСписка:
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
                                paragraph.Format.SpaceBefore != 0) && (i != 0 && infos[i - 1].Type != ParagraphInfo.ParagraphType.Таблица))
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
                    case ParagraphInfo.ParagraphType.БиблиографическоеОписаниеИсточника:
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
                                paragraph.Format.SpaceBefore != 0) && (i != 0 && infos[i - 1].Type != ParagraphInfo.ParagraphType.Таблица))
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
                    case ParagraphInfo.ParagraphType.ЗаголовокСпискаЛитературы:
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
                if (infos[i].Type == ParagraphInfo.ParagraphType.Текст
                    || infos[i].Type == ParagraphInfo.ParagraphType.ЭлементНумерованногоСписка
                    || infos[i].Type == ParagraphInfo.ParagraphType.ЭлементМаркерованногоСписка)
                {
                    string text = document.Paragraphs[infos[i].Index].Range.Text.Trim().ToLower();

                    for (int j = 0; j < references.Count; j++)
                    {
                        if (!references[j].HasRef && references[j].Index > infos[i].Index)
                        {
                            string objectname = "объект";
                            switch (references[j].Type)
                            {
                                case ParagraphInfo.ParagraphType.ПодрисуночнаяПодпись:
                                    objectname = ObjectTitleMarker.FigRef;
                                    break;
                                case ParagraphInfo.ParagraphType.НазваниеТаблицы:
                                    objectname = ObjectTitleMarker.TableRef;
                                    break;
                                case ParagraphInfo.ParagraphType.НазваниеЛистинга:
                                    objectname = ObjectTitleMarker.CodeRef;
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
                        case ParagraphInfo.ParagraphType.ПодрисуночнаяПодпись:
                            objectname = ObjectTitleMarker.FigTitle;
                            break;
                        case ParagraphInfo.ParagraphType.НазваниеТаблицы:
                            objectname = ObjectTitleMarker.TableTitle;
                            break;
                        case ParagraphInfo.ParagraphType.НазваниеЛистинга:
                            objectname = ObjectTitleMarker.CodeTitle;
                            break;
                    }

                    infos[i].Problems.Add($"отсутствует ссылка на {objectname} " + infos[i].NumberObjectInText + " перед объектом");
                }
            }


            string summaryComment = "Общие замечания по документу:";

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
                document.Paragraphs[1].Range.Comments.Add(document.Paragraphs[1].Range.Words[1], summaryComment);
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
                        document.Comments.Add(firstLiteralRangeComment, comment);
                    }

                }
               catch { } // есть скрытые пустые абзацы в таблицах - на них не ставит примечания, ирод
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
