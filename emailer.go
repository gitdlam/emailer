package main

import (
	"io/ioutil"
	"log"
	"sort"
	"strings"
	"sync"
	"time"
	//	"bytes"
	"fmt"
	"os/exec"
	"path/filepath"
	//	"io/ioutil"
	//	"log"
	"net/http"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/gorhill/cronexpr"
	//	"github.com/gitdlam/common"
	"github.com/jordan-wright/email"
)

type crontabType struct {
	minute      string
	hour        string
	dayOfMonth  string
	month       string
	dayOfWeek   string
	translation string
	cmd         string
	output      string
	recipients  string
	emailFrom   string
	subject     string
	body        string
}

type entryType struct {
	planned     time.Time
	translation string
	cmd         string
	output      string
	recipients  string
	emailFrom   string
	subject     string
	body        string
	done        bool
}

var entries []entryType
var entriesLock sync.RWMutex

func pingResponse(w http.ResponseWriter, req *http.Request) {
	fmt.Fprintf(w, appName())

}

func folderResponse(w http.ResponseWriter, req *http.Request) {
	fmt.Fprintf(w, appFolder())

}

func listResponse(w http.ResponseWriter, req *http.Request) {
	entriesLock.Lock()
	defer entriesLock.Unlock()
	list := `<html><head><style> table a:link {
	color: #666;
	font-weight: bold;
	text-decoration:none;
}
table a:visited {
	color: #999999;
	font-weight:bold;
	text-decoration:none;
}
table a:active,
table a:hover {
	color: #bd5a35;
	text-decoration:underline;
}
table {
	font-family:Arial, Helvetica, sans-serif;
	color:#222;
	font-size:14px;
	text-shadow: 1px 1px 0px #fff;
	background:#eaebec;
	margin:10px;
	border:#ccc 1px solid;

	-moz-border-radius:3px;
	-webkit-border-radius:3px;
	border-radius:3px;

	-moz-box-shadow: 0 1px 2px #d1d1d1;
	-webkit-box-shadow: 0 1px 2px #d1d1d1;
	box-shadow: 0 1px 2px #d1d1d1;
}
table th {
	padding:21px 25px 22px 25px;
	border-top:1px solid #fafafa;
	border-bottom:1px solid #e0e0e0;

	background: #ededed;
	background: -webkit-gradient(linear, left top, left bottom, from(#ededed), to(#ebebeb));
	background: -moz-linear-gradient(top,  #ededed,  #ebebeb);
}
table th:first-child {
	text-align: left;
	padding-left:20px;
}
table tr:first-child th:first-child {
	-moz-border-radius-topleft:3px;
	-webkit-border-top-left-radius:3px;
	border-top-left-radius:3px;
}
table tr:first-child th:last-child {
	-moz-border-radius-topright:3px;
	-webkit-border-top-right-radius:3px;
	border-top-right-radius:3px;
}
table tr {
	text-align: center;
	padding-left:20px;
}
table td:first-child {
	text-align: left;
	padding-left:20px;
	border-left: 0;
}
table td {
	padding:18px;
	border-top: 1px solid #ffffff;
	border-bottom:1px solid #e0e0e0;
	border-left: 1px solid #e0e0e0;

	background: #fafafa;
	background: -webkit-gradient(linear, left top, left bottom, from(#fbfbfb), to(#fafafa));
	background: -moz-linear-gradient(top,  #fbfbfb,  #fafafa);
}
table tr.even td {
	background: #f6f6f6;
	background: -webkit-gradient(linear, left top, left bottom, from(#f8f8f8), to(#f6f6f6));
	background: -moz-linear-gradient(top,  #f8f8f8,  #f6f6f6);
}
table tr:last-child td {
	border-bottom:0;
}
table tr:last-child td:first-child {
	-moz-border-radius-bottomleft:3px;
	-webkit-border-bottom-left-radius:3px;
	border-bottom-left-radius:3px;
}
table tr:last-child td:last-child {
	-moz-border-radius-bottomright:3px;
	-webkit-border-bottom-right-radius:3px;
	border-bottom-right-radius:3px;
}
table tr:hover td {
	background: #f2f2f2;
	background: -webkit-gradient(linear, left top, left bottom, from(#f2f2f2), to(#f0f0f0));
	background: -moz-linear-gradient(top,  #f2f2f2,  #f0f0f0);	
}</style></head><body><table><thead><tr><th>Planned</th><th>Timing</th><th>Output</th><th>Recipients</th></tr></thead>`

	for _, e := range entries {
		list = list + "\n<tr><td>" + e.planned.Format("2006-01-02 15:04") + "</td><td>" + e.translation + "</td><td>" + e.output + "</td><td>" + e.recipients + "</td></tr>"
	}
	list = list + "</table></body></html>"
	fmt.Fprintf(w, list)

}

func appName() string {
	ex, err := os.Executable()
	if err != nil {
		panic(err)
	}
	_, name := filepath.Split(ex)

	return strings.Split(name, ".exe")[0]
}

func appFolder() string {
	ex, err := os.Executable()
	if err != nil {
		panic(err)
	}
	folder, _ := filepath.Split(ex)

	return folder
}

func HTTPServe() {

	http.HandleFunc("/ping", pingResponse)
	http.HandleFunc("/folder", folderResponse)
	http.HandleFunc("/list", listResponse)

	http.ListenAndServe(":1777", nil)

}

func getCronTabs() []crontabType {
	var crontabs []crontabType
	xlsxFile, err := excelize.OpenFile(appFolder() + "emailer.xlsx")

	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}

	rows := xlsxFile.GetRows("Sheet1")
	for i, _ := range rows {
		if i < 2 {
			continue
		}
		if rows[i][0] == "" {
			break
		}
		crontab := crontabType{}
		crontab.minute = rows[i][0]

		crontab.hour = rows[i][1]
		crontab.dayOfMonth = rows[i][2]
		crontab.month = rows[i][3]
		crontab.dayOfWeek = rows[i][4]
		crontab.translation = rows[i][5]
		crontab.cmd = rows[i][6]
		crontab.output = rows[i][7]
		crontab.recipients = rows[i][8]
		crontab.emailFrom = rows[i][9]
		crontab.subject = rows[i][10]
		crontab.body = rows[i][11]
		crontabs = append(crontabs, crontab)
		//		log.Println(crontab.minute, crontab.hour, crontab.dayOfMonth, crontab.month, crontab.dayOfWeek, crontab.translation, crontab.cmd, crontab.output, crontab.emails)

	}
	return crontabs

}

func main() {
	refreshSchedule()

	go monitorSchedule()

	HTTPServe()

}

func refreshSchedule() {
	entriesLock.Lock()
	var crontabs []crontabType
	entries = nil
	crontabs = getCronTabs()

	for _, c := range crontabs {
		nextTime := time.Now()
		expr := cronexpr.MustParse(strings.Join([]string{c.minute, c.hour, c.dayOfMonth, c.month, c.dayOfWeek}, " "))
		for nextTime.Before(time.Now().AddDate(0, 0, 60)) {
			nextTime = expr.Next(nextTime)
			//			log.Println(nextTime)
			//			time.Sleep(time.Second)
			if nextTime.IsZero() || nextTime.After(time.Now().AddDate(0, 0, 60)) {
				break
			}

			entries = append(entries, entryType{planned: nextTime, translation: c.translation, cmd: c.cmd, output: c.output, recipients: c.recipients, emailFrom: c.emailFrom, subject: c.subject, body: c.body})
		}

	}
	expr := cronexpr.MustParse(strings.Join([]string{"10", "3", "*", "*", "*"}, " "))
	entries = append(entries, entryType{planned: expr.Next(time.Now().Add(time.Minute)), translation: "Refresh schedule from config file", subject: "refresh_schedule"})
	sort.SliceStable(entries, func(i, j int) bool {
		return entries[i].planned.Format("200601021504") < entries[j].planned.Format("200601021504")
	})
	entriesLock.Unlock()
}

func monitorSchedule() {
	var refresh bool
	for {
		entriesLock.Lock()
		for i, e := range entries {
			if !e.done && time.Now().After(e.planned) && e.subject != "refresh_schedule" {
				cmd := exec.Command("cmd", "/c", e.cmd)
				// log.Println("start cmd")
				cmd.Run()
				// log.Println("finished cmd")
				sendEmail(e.emailFrom, e.recipients, e.subject, e.body, e.output)
				// log.Println("finished email")
				entries[i].done = true
			}
			if time.Now().After(e.planned) && e.subject == "refresh_schedule" {
				refresh = true
			}
		}
		entriesLock.Unlock()

		if refresh {
			refreshSchedule()
			refresh = false
		} else {
			time.Sleep(time.Second * 15)
		}

	}
}

func sendEmail(from string, to string, subject string, body string, attachment string) {
	e := email.NewEmail()
	e.From = from
	e.To = strings.Split(to, ";")
	e.Subject = subject
	html := readHTML(body)
	if html == nil {
		e.HTML = []byte("<html><body>Please find report attached.</body></html>")
	} else {
		e.HTML = html
	}
	e.AttachFile(attachment)
	err := e.Send("smtp:25", nil)
	if err != nil {
		log.Printf("%s\n", err)
	}

}

func readHTML(fileName string) []byte {
	data, err := ioutil.ReadFile(fileName)
	if err != nil {
		return nil
	}
	return data
}
