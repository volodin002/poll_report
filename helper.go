package main

import (
	"strconv"
)

// col and row NOT zero-based
func cellName(col, row int) string {
	res := ""
	for col > 0 {
		rem := col%26
		if rem == 0 {
			res = string('Z') + res 
			col = (col/26) - 1
		} else {
			// 65 - ASCII Uppercase A
			res = string(65 + (rem-1)) + res 
			col = col/26
		}
	}

	res = res + strconv.Itoa(row)

	return res
}

// return NOT zero-based column number
func colNumber(col string) int  {
	runes := []rune(col)
	ln := len(runes)
	sum := 0
	for i := 0; i < ln; i++ {
		sum *= 26;
		// 65 - ASCII Uppercase A
        sum += (int(runes[i]) - 65 + 1);
    }

    return sum;
}