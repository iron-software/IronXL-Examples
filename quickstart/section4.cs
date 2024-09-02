// Set cell's value and styles
workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double;