# Excel SUMIF & SUMIFS Questions For Data Analyst

This project develops understanding practical use of Excel `SUMIF` and `SUMIFS` through real-world, interview-level analytical questions.

---

## SUMIF & SUMIFS Practice Questions & Answers

---

### 1. Find total sales from Lahore
```excel
=SUMIF(C2:C31,"Lahore",E2:E31)
```
### 2. Find total sales of "Laptop" products
```excel
=SUMIF(D2:D31,"Laptop",E2:E31)
```
### 3. Find total sales from Karachi where product is "Laptop"
```excel
=SUMIFS(E2:E31,C2:C31,"Karachi",D2:D31,"Laptop")
```
### 4. Find total sales where Region = Lahore and Sales > 500
```excel
=SUMIFS(E2:E31,C2:C31,"Lahore",E2:E31,">500")
```
### 5. Find total sales where Product = "Mobile" and Sales < 300
```excel
=SUMIFS(E2:E31,D2:D31,"Mobile",E2:E31,"<300")
```
### 6. Find total sales where Product = "Laptop" and Region is NOT Karachi
```excel
=SUMIFS(E2:E31,D2:D31,"Laptop",C2:C31,"<>Karachi")
```
### 7. Find total sales where Region = Lahore, Product = "Laptop", and Sales > 800
```excel
=SUMIFS(E2:E31,C2:C31,"Lahore",D2:D31,"Laptop",E2:E31,">800")
```
### 8. Find total sales where Sales are between 300 and 800
```excel
=SUMIFS(E2:E31,E2:E31,">=300",E2:E31,"<=800")
```
### 9. Find total sales where SalesPerson name starts with "A"
```excel
=SUMIF(B2:B31,"A*",E2:E31)
```

### 10. Find total sales where Region is NOT Islamabad, Product = "Tablet", and Sales > 400
```excel
=SUMIFS(E2:E31,C2:C31,"<>Islamabad",D2:D31,"Tablet",E2:E31,">400")
```
### 11. Find total sales where SalesPerson = "Ali", Region = Lahore, and Product is NOT "Tablet"
```excel
=SUMIFS(E2:E31,B2:B31,"Ali",C2:C31,"Lahore",D2:D31,"<>Tablet")
```

### Points to remember:

1. SUMIF = single condition summing <br>
2. SUMIFS = multiple conditions summing <br>
3. Use ">" "<" ">=" "<=" for numeric conditions <br>
4. Use "<>" for NOT condition <br>
5. Wildcards: * (multiple characters), ? (single character) <br>
