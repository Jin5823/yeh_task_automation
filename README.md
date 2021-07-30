# 업무 자동화
> Work automation, 工作流程自动化

## 요약

무명 음악가 [예현](https://www.instagram.com/itsyehworld)은 스폰서들의 후원을 받아 [MV영상](https://www.youtube.com/watch?v=CLdzzVFq33c)을 제작했다. 
스폰서들에게 받은 후원의 사용내역을 경비 보고서로 작성하고 각 스폰서들에게 전달하는 업무를 해야하는데, 시간이 많이 소요되는 루즈한 작업이기에 
자동화를 통해 시간 소요를 최소화하고 결과도 빠르게 전달한다.  

## 요소

- file

|파일명|설명|
|:-|:-|
|`expense_book.xlsx`|지출내역에대한 정보를 담고 있는 파일|
|`sponsor_book.xlsx`|스폰서들의 정보 목록|
|`print_style.xlsx`|전달할 경비 보고서의 양식 파일|
|`Expense_Report.py`|데이터를 받아 경비 보고서를 자동으로 작성하고 압축하여 저장|

- flow

사전에 데이터를 기록한 파일에서 데이터를 하나식 가져와 경비 보고서 양식 파일에 작성한다. 양식 파일은 엑셀 형태의 파일이기에 파이썬의 `openpyxl`패키지를 사용하고, 
작성 완료된 엑셀 파일은 보기 편한 [pdf 파일](https://github.com/Jin5823/yeh_task_automation/blob/master/result_pdf)로 `win32com`패키지를 사용해 전환한다. 이어서 `subprocess`패키지를 사용해 cmd 명령을 호출하고, pdf파일을 오픈소스인 7zip을 
이용해 [zip 파일](https://github.com/Jin5823/yeh_task_automation/tree/master/result_zip)로 압축하며, 비밀번호는 스폰서의 전화번호 혹은 아이디의 뒷자리 4개로 설정한다. 

- result

<img src="https://raw.githubusercontent.com/Jin5823/Git-Test/master/src/img_11.JPG" />

## 결론 및 토론

모든 후원자의 zip파일은 [Onedrive](https://1drv.ms/f/s!Aos6j-DPzfAzmjgw_MWYBX8_2Ns6)에 저장하여 오픈링크를 생성하고, 링크를 타고 누구나 자신의 후원 목록과 사용내역을 조회할 수 있으며, 
암호화했기에 자신외에 다른 사람들은 후원금액과 내역을 확일할 수 없다. 일일이 양식에 맞게 내역을 옮겨서 파일을 작성하는 루즈한 업무는 100줄도 안되는 코드에 간단하고 빠르게 
결과를 얻을 수 있어 만족스러웠다.


