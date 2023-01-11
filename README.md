<h4>VBA_data_separation</h4>
<p>Concepts for data separation in VBA (separate, check data, store to multiple workbooks, etc.)</p>
<ol>
<li>Concept for separating a year's data into four separate Worksheets for each quarter of the year. The data contains a column which designates the quarter of the year. After that, the four Worksheets are stored into separate Workbooks. Each Workbook contains one Worksheet (separate_data_from_wks_to_separate_wkbs.bas).</li>
<li>Concept for the decipherment of combinations which are oriented column by column with a lookup table. Columns of the sample table have the following structure: A1234/B2345/E1711,A1234/D3456,B2345/C8976,E1711/A1234,E1711/C8976,C8976
The lookup table holds the decipherment for the codes: column1 = code, column2 = decipherment
{'A1234': 'alpha', 'B2345':'beta', 'C8976':'gamma', 'D3456':'delta', 'E1711':'epsilon'}
The program creates the decipherment for the column in the sample table:
alpha/beta/epsilon,alpha/delta,beta/gamma,epsilon/alpha,epsilon/gamma,gamma</li>
</ol>
