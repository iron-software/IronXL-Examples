// Set a formula
workSheet["A1"].Value = "=SUM(A2:A10)";

// Get the calculated value
decimal sum = workSheet["A1"].DecimalValue;