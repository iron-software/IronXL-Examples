// Use IronXL built-in aggregations
decimal sum = workSheet["A2:A11"].Sum();
decimal avg = workSheet["B2:B11"].Avg();
decimal max = workSheet["C2:C11"].Max();
decimal min = workSheet["D2:D11"].Min();

// Assign value to cells
workSheet["A12"].Value = sum;
workSheet["B12"].Value = avg;
workSheet["C12"].Value = max;
workSheet["D12"].Value = min;