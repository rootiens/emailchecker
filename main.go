package main

import (
	"fmt"
	"strconv"
	"strings"
	"sync"

	emailverifier "github.com/AfterShip/email-verifier"
	excelize "github.com/xuri/excelize/v2"
)

var (
	verifier = emailverifier.
		NewVerifier().
		EnableSMTPCheck()
)

func checkMail(sourcemail string, wg *sync.WaitGroup, mu *sync.Mutex, results map[string]bool) {
	defer wg.Done()

	ret, err := verifier.CheckSMTP(strings.Split(sourcemail, "@")[1], strings.Split(sourcemail, "@")[0])

	mu.Lock()
	defer mu.Unlock()

	if err != nil {
		results[sourcemail] = false
		return
	}

	if ret.HostExists && !ret.Disabled && ret.Deliverable {
		results[sourcemail] = true
		return
	}

	results[sourcemail] = false
}

func main() {
    fmt.Print("Enter filename: ")
    var input string
    fmt.Scanln(&input)

	file, err := excelize.OpenFile(input)
	if err != nil {
		panic(err)
	}

	defer func() {
		if err := file.Close(); err != nil {
			panic(err)
		}
	}()

	rows, err := file.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	var wg sync.WaitGroup
	var mu sync.Mutex
	results := make(map[string]bool)

	for _, row := range rows {
		for _, colCell := range row {
			wg.Add(1)
			go checkMail(colCell, &wg, &mu, results)
		}
	}

	wg.Wait()

	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	index, err := f.NewSheet("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	idx := 1

	for email, ok := range results {
		if ok {
			f.SetCellValue("Sheet1", "A"+strconv.Itoa(idx), email)
            idx += 1
		}
	}

	f.SetActiveSheet(index)
	if err := f.SaveAs("VerifiedMails.xlsx"); err != nil {
		fmt.Println(err)
	}
}
