import requests
from bs4 import BeautifulSoup 
from openpyxl import Workbook
import re
from openpyxl import load_workbook
import openpyxl
import os.path

def main():
    file = 'lotto.xlsx'
    if os.path.isfile(file):
        new_ball()
        lastrow()
        if int(keyword[6]) == last_row -1:
            print('최신버전입니다.')
        else:
            wb = openpyxl.load_workbook('lotto.xlsx')
            sheet = wb.active
            row = sheet.max_row
            print(row)
            print('업데이트를 시작합니다.')
            for j in range(row,int(keyword[6])+1):
                basic_url = "https://www.dhlottery.co.kr/gameResult.do?method=byWin&drwNo=" 
                resp = requests.get(basic_url + str(j)) 
                soup = BeautifulSoup(resp.text, "lxml") 
                line = str(soup.find("meta", {"id" : "desc", "name" : "description"})['content']) 

                begin = line.find("당첨번호")
                begin = line.find(" ", begin) + 1 
                end = line.find(".", begin) 
                numbers = line[begin:end] 
                print("당첨번호" + str(j) +"회" , numbers)
                split_num = re.split('[, +]',numbers)
                write_ws = wb.active
                write_ws['A'+str(j+1)] = str(j) + "회"
                
                write_ws['B'+str(j+1)] = int(split_num[0])
                write_ws['C'+str(j+1)] = int(split_num[1])
                write_ws['D'+str(j+1)] = int(split_num[2])
                write_ws['E'+str(j+1)] = int(split_num[3])
                write_ws['F'+str(j+1)] = int(split_num[4])
                write_ws['G'+str(j+1)] = int(split_num[5])
                write_ws['H'+str(j+1)] = int(split_num[6])

            write_ws['H1'] = "추가번호"   
            write_ws['B1'] = "1번째 번호"
            write_ws['C1'] = "2번째 번호"
            write_ws['D1'] = "3번째 번호"
            write_ws['E1'] = "4번째 번호"
            write_ws['F1'] = "5번째 번호"
            write_ws['G1'] = "6번째 번호"

            wb.save('lotto.xlsx')


            print('업데이트 완료')

    else:
        dowmload()

    
def dowmload():
    makefile()
    new_ball()
    lastrow()
    print("다운로드 시작합니다.")
    basic_url = "https://www.dhlottery.co.kr/gameResult.do?method=byWin&drwNo=" 
    write_wb = Workbook()
    for i in range(1,int(keyword[6])+1):
        resp = requests.get(basic_url + str(i)) 
        soup = BeautifulSoup(resp.text, "lxml") 
        line = str(soup.find("meta", {"id" : "desc", "name" : "description"})['content']) 

        begin = line.find("당첨번호")
        begin = line.find(" ", begin) + 1 
        end = line.find(".", begin) 
        numbers = line[begin:end] 
        print("당첨번호" + str(i) +"회" , numbers)
        split_num = re.split('[, +]',numbers)
        write_ws = write_wb.active
        if i == 0:
            pass
        else:
            write_ws['A'+str(i+1)] = str(i) + "회"

            write_ws['B'+str(i+1)] = int(split_num[0])
            write_ws['C'+str(i+1)] = int(split_num[1])
            write_ws['D'+str(i+1)] = int(split_num[2])
            write_ws['E'+str(i+1)] = int(split_num[3])
            write_ws['F'+str(i+1)] = int(split_num[4])
            write_ws['G'+str(i+1)] = int(split_num[5])
            write_ws['H'+str(i+1)] = int(split_num[6])
    write_ws['H1'] = "추가번호"   
    write_ws['B1'] = "1번째 번호"
    write_ws['C1'] = "2번째 번호"
    write_ws['D1'] = "3번째 번호"
    write_ws['E1'] = "4번째 번호"
    write_ws['F1'] = "5번째 번호"
    write_ws['G1'] = "6번째 번호"

    write_wb.save('lotto.xlsx')

    print('다운로드 완료')



def makefile():
    write_wb = Workbook()
    write_wb.save('lotto.xlsx')

def lastrow():
    global last_row
    wb = openpyxl.load_workbook('lotto.xlsx')
    sheet = wb.active
    last_row = sheet.max_row
    # print(last_row-1)
    # last_col_a_value = sheet.cell(column=2, row=last_row+1).value = "asdasdad"
    # wb.save('lotto.xlsx')

def new_ball():
    global keyword
    url = 'https://www.nlotto.co.kr/common.do?method=main'
    a = requests.get(url)
    soup = BeautifulSoup(a.text, "lxml")
    find = soup.find(id="lottoDrwNo")
    keyword = re.split('[< = " >]',str(find))
    # print(keyword)
    # for i in range(1,6):
    #     ball_find = soup.find(class_="ball_645 lrg ball{}".format(i))
    #     keyword = re.split('[< = " >]',str(ball_find))
    #     print(keyword[12])
    # bonus_ball = soup.find(id="bnusNo")
    # print(bonus_ball)


if __name__ == "__main__":
    main()