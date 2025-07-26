# Missing Functions Report

This report compares the functions listed in `/docs/function-implementation-checklist.md` (marked with 📝) against the functions actually implemented in `/src/data/individualFunctionTests.ts`.

## Summary

- **Functions marked with 📝 in checklist**: 265
- **Functions found in individualFunctionTests.ts**: 385 total (but only 179 match the checklist)
- **Missing functions**: 86

## Missing Functions by Category

### 1. Basic Operators (10 functions - ALL MISSING)
- ADD_OPERATOR
- SUBTRACT_OPERATOR
- MULTIPLY_OPERATOR
- DIVIDE_OPERATOR
- GREATER_THAN
- LESS_THAN
- GREATER_THAN_OR_EQUAL
- LESS_THAN_OR_EQUAL
- EQUAL
- NOT_EQUAL

### 2. Financial Functions (29 functions)
- ACCRINT
- ACCRINTM
- AMORDEGRC
- AMORLINC
- COUPDAYBS
- COUPDAYS
- COUPDAYSNC
- COUPNCD
- COUPNUM
- COUPPCD
- CUMIPMT
- CUMPRINC
- DISC
- DOLLARDE
- DOLLARFR
- DURATION
- EFFECT
- INTRATE
- MDURATION
- NOMINAL
- ODDFPRICE
- ODDFYIELD
- ODDLPRICE
- ODDLYIELD
- PRICE
- PRICEDISC
- PRICEMAT
- RECEIVED
- TBILLEQ
- TBILLPRICE
- TBILLYIELD
- YIELD
- YIELDDISC
- YIELDMAT

### 3. Engineering Functions (14 functions)
- BESSELI
- BESSELJ
- BESSELK
- BESSELY
- BITAND
- BITLSHIFT
- BITOR
- BITRSHIFT
- BITXOR
- DELTA
- GESTEP
- IMCOSH
- IMCOT
- IMCSC
- IMCSCH
- IMSEC
- IMSECH
- IMSINH

### 4. Cube Functions (7 functions)
- CUBEKPIMEMBER
- CUBEMEMBER
- CUBEMEMBERPROPERTY
- CUBERANKEDMEMBER
- CUBESET
- CUBESETCOUNT
- CUBEVALUE

### 5. Web/Import Functions (8 functions)
- IMAGE
- IMPORTDATA
- IMPORTFEED
- IMPORTHTML
- IMPORTRANGE
- IMPORTXML

### 6. Google Sheets Functions (6 functions)
- DETECTLANGUAGE
- GOOGLEFINANCE
- GOOGLETRANSLATE
- TO_DATE
- TO_DOLLARS
- TO_PERCENT
- TO_TEXT

### 7. Other Functions (3 functions)
- GPT
- ISBETWEEN
- ISOMITTED
- STOCKHISTORY

## Critical Findings

1. **All basic operators are missing** - This is surprising as these are fundamental operations (+, -, *, /, comparison operators)

2. **Most financial functions are missing** - A significant portion of advanced financial calculations are not in the individual tests

3. **All cube functions are missing** - OLAP cube-related functions are completely absent

4. **Most web/import functions are missing** - Data import capabilities are largely untested

## Recommendation

The checklist claims that all 507 functions have individual tests, but this analysis shows 86 functions from the checklist are missing from individualFunctionTests.ts. Priority should be given to:

1. Basic operators (critical)
2. Commonly used financial functions
3. Web/import functions for data integration

Note: There may be a discrepancy between how functions are named in the checklist vs. the implementation file, or some functions might be tested differently.