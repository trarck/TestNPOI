using UnityEngine;
using System.Collections;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

public class TestRead : MonoBehaviour {

	// Use this for initialization
	void Start ()
    {
        Debug.Log(Application.dataPath + "/Resources/Data/Test.xlsx");

        if(File.Exists(Application.dataPath + "/Resources/Data/Test.xlsx"))
        {
            Debug.Log("OK");
        }
        else
        {
            Debug.Log("No file");
        }

        XSSFWorkbook workbook = new XSSFWorkbook(Application.dataPath+ "/Resources/Data/Test.xlsx");
        ISheet testSheet = workbook.GetSheet("Test");
        IEnumerator iter = testSheet.GetRowEnumerator();
        while (iter.MoveNext())
        {
            IRow row = iter.Current as IRow;
            string line = "";
            foreach(ICell cell in row.Cells)
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        line += cell.StringCellValue;
                        break;
                    case CellType.Numeric:
                        line += cell.NumericCellValue;
                        break;
                    case CellType.Boolean:
                        line += cell.BooleanCellValue;
                        break;
                    case CellType.Error:
                        line += cell.ErrorCellValue;
                        break;
                }

                line += "[" + cell.CellType + "]"+ (cell.CellComment!=null ? ("("+cell.CellComment.String.String+")"):"")+"\t";
                
            }

            Debug.Log(line);
        }

    }
	
	// Update is called once per frame
	void Update () {
	
	}
}
