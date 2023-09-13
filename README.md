# VBA-Challenge
Module 2 Challenge - VBA

Added results & VBA script

Received code from Learning Assistant for Match + Max Functions
row_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
ticket_max = Cells(row_number + 1, "put_the_column_number_here_from_excel")
