using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SamuraiTools.OpenXml.Spreadsheet
{
    public sealed class WorksheetCollection : ICollection<Worksheet>
    {
        private SpreadsheetDocument document;

        public WorksheetCollection(SpreadsheetDocument document) { this.document = document; }

        public Worksheet this[string name]
        {
            get
            {
                Sheet sheet = document.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => { return s.Name == name; });
                if (sheet == null)
                {
                    return null;
                }
                else
                {
                    return ((WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                }
            }
        }

        public Worksheet this[int index]
        {
            get
            {
                return ((WorksheetPart)document.WorkbookPart.GetPartById(document.WorkbookPart.Workbook.Descendants<Sheet>().ElementAt(index).Id)).Worksheet;
            }
        }

        public int Count => document.WorkbookPart.WorksheetParts.Count();

        public bool IsReadOnly => false;

        /// <summary>
        /// Add this Worksheet to the collection with the provided name.
        /// </summary>
        /// <param name="item">The Worksheet to add.</param>
        /// <param name="name">The name for this Worksheet.</param>
        public void Add(Worksheet item, string name)
        {
            string trimmedName = name?.Trim();

            // Add a blank WorksheetPart.
            WorksheetPart newWorksheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = item;

            Sheets sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = document.WorkbookPart.GetIdOfPart(newWorksheetPart);

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Append the new worksheet and associate it with the workbook.
            sheets.Append(new Sheet() { Id = relationshipId, SheetId = sheetId, Name = string.IsNullOrEmpty(trimmedName) ? "Sheet" + sheetId.ToString() : trimmedName });
        }

        /// <summary>
        /// Add this Worksheet to the collection. The name will be generated automatically.
        /// </summary>
        /// <param name="item"></param>
        public void Add(Worksheet item)
        {
            Add(item, null);
        }

        /// <summary>
        /// Add a new Worksheet with empty SheetData to the collection with the provided name.
        /// </summary>
        /// <param name="name">The name for the new Worksheet.</param>
        /// <returns></returns>
        /// <remarks>Name can be null or empty, in which case a name will be generated automatically.</remarks>
        public Worksheet AddNew(string name)
        {
            Add(new Worksheet(new SheetData()), name);

            return this[Count - 1];
        }

        /// <summary>
        /// Clear all Worksheets from the collection.
        /// </summary>
        public void Clear()
        {
            foreach (var part in document.WorkbookPart.WorksheetParts.ToList())
            {
                string relationshipId = document.WorkbookPart.GetIdOfPart(part);
                document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Single(s => s.Id == relationshipId).Remove();
            }

            document.WorkbookPart.DeleteParts(document.WorkbookPart.WorksheetParts);
        }

        /// <summary>
        /// Whether the provided Worksheet is in the collection.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public bool Contains(Worksheet item)
        {
            return document.WorkbookPart.WorksheetParts.Any(x => x.Worksheet == item);
        }

        /// <summary>
        /// Whether a Worksheet with the provided name is in the collection.
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public bool Contains(string name)
        {
            return document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Any(s => s.Name == name);
        }

        /// <summary>
        /// Copy the Worksheets in the collection to an array, starting at the provided index.
        /// </summary>
        /// <param name="array"></param>
        /// <param name="arrayIndex"></param>
        public void CopyTo(Worksheet[] array, int arrayIndex)
        {
            foreach (var part in document.WorkbookPart.WorksheetParts)
            {
                array[arrayIndex++] = part.Worksheet;
            }
        }

        /// <summary>
        /// Remove the Worksheet with the provided name from the collection.
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public bool Remove(string name)
        {
            Sheet sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().FirstOrDefault(s => s.Name == name);

            if (sheet == null)
            {
                return false;
            }

            string relationshipId = sheet.Id;
            sheet.Remove();
            return document.WorkbookPart.DeletePart(relationshipId);
        }

        /// <summary>
        /// Remove the provided Worksheet from the collection.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public bool Remove(Worksheet item)
        {
            WorksheetPart part = document.WorkbookPart.WorksheetParts.FirstOrDefault(x => x.Worksheet == item);

            if (part == null)
            {
                return false;
            }

            string relationshipId = document.WorkbookPart.GetIdOfPart(part);
            document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Single(s => s.Id == relationshipId).Remove();

            return document.WorkbookPart.DeletePart(part);
        }

        public IEnumerator<Worksheet> GetEnumerator()
        {
            //using ToList because otherwise any addition to the WorkbookPart results in error.
            foreach (var worksheet in document.WorkbookPart.WorksheetParts.Select(part => part.Worksheet).ToList())
            {
                yield return worksheet;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}
