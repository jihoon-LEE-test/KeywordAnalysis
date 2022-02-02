##재료는 targetlist 엑셀파일과 oursite_all 파일 입니다.
## 실행시키면 3개나라(저지, 맨섬, 건지)의 결과가 타겟리스트 파일에 값이 입력되고 3개의 나라에대한 각 엑셀파일이 저장됩니다.
# @korea.ac.kr
import os
import openpyxl
import pandas as pd
import warnings
import requests

#경고 끄기
warnings.filterwarnings("ignore")

## 모든 열을 출력한다.
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

kws_ko = ['하나님의 교회', '하나님의교회 세계복음선교협회', '안상홍', '어머니 하나님', '장길자']
kws_en = ['Church of God', 'World Mission Society Church of God', 'Ahnsahnghong', 'God the Mother', 'Zahng Gil-jah']
kws_es = ['Iglesia de Dios', 'Iglesia de Dios sociedad misionera mundial', 'Ahnsahnghong', 'Dios Madre', 'Zahng Gil-jah']


params1 = {
    'api_key': '38CBEF8539454506B1DC3FED84038381',
    'q': 'World Mission Society Church of God',
    'num' : '50',
    'hl' :"ko",
    'gl': "kr",
    'domain' : 'google.com'
}

params2 = {
  'api_key': '38CBEF8539454506B1DC3FED84038381',
    'q': 'World Mission Society Church of God',
    'num' : '50',
    'hl' : "en",
    'gl': "us",
    'domain' : 'google.com',
}

params3 = {
  'api_key': '38CBEF8539454506B1DC3FED84038381',
    'q': 'World Mission Society Church of God',
    'num' : '50',
    'hl' : "es",
    'gl': "pe",
    'domain' : 'google.com'
}

# 응답받은 데이터(json)를 판다스 데이터프레임 변경 후, 엑셀 형태로 저장. 각 키워드로 시트명 생성
def json_to_df(f):
    df = pd.DataFrame(f, dtype='unicode')
    return df


# 무시
pd.set_option('mode.chained_assignment',  None) # <==== 경고를 끈다

# DataFrame을 엑셀 각 시트로 출력하기
def df_to_excel_each_sheet(df, kw, f):
    # 최초 생성 이후 mode는 append; 새로운 시트를 추가합니다.
    if not os.path.exists(f):
        with pd.ExcelWriter(f, mode='w', engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=kw)
    else:
        with pd.ExcelWriter(f, mode='a', engine='openpyxl', if_sheet_exists="replace") as writer:
            df.to_excel(writer, index=False, sheet_name=kw)

def df_to_excel_each_sheet(df, kw, f):
    # 최초 생성 이후 mode는 append; 새로운 시트를 추가합니다.
    if not os.path.exists(f):
      with pd.ExcelWriter(f, mode='w', engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=kw)
    else:
      with pd.ExcelWriter(f, mode='a', engine='openpyxl', if_sheet_exists="replace") as writer:
        df.to_excel(writer, index=False, sheet_name=kw)

def do_api(param):
    api_result1 = requests.get('https://api.serpwow.com/search', params1)
    print("ko 완료")
    api_result2 = requests.get('https://api.serpwow.com/search', params2)
    print("us 완료")
    api_result3 = requests.get('https://api.serpwow.com/search', params3)
    print("es 완료")

    data1 = api_result1.json()
    data2 = api_result2.json()
    data3 = api_result3.json()

    organic_results1 = data1['organic_results']
    organic_results2 = data2['organic_results']
    organic_results3 = data3['organic_results']

    organic_results_df1 = json_to_df(organic_results1)
    organic_results_df2 = json_to_df(organic_results2)
    organic_results_df3 = json_to_df(organic_results3)

    #################### 파일 리딩 파트 ########################
    df3_read = pd.read_excel('our_site_all.xlsx', sheet_name=0)

    # 카운트용 변수들
    heat_count = 0
    print('시작!')
    # #################엑셀 리딩 -> 사이트 구분 part
    #데이터 프레임을 받아서 하나하나 따로 값을 요리한 후에 값 저장하기 위함.


def get_occu_rate(s_result, kw,f_name, excel_p):

    # 키워드 데이터 불러들이기
    df1_read = s_result

    #1~3열까지만 불러와서 데이터를 다룬다.1. 포지션, 2 타이틀, 3. 링크. 4열에 링크2를 추가하는 이유는 나중에 원복하기 위함.
    df2_read = df1_read.iloc[:, :]
    # 주소 데이터 전처리
    df2_read['link2'] = df2_read['link']
    df2_read['link'].str.lower()
    df2_read['link'] = df2_read['link'].str.replace("http://", " ", regex = False)
    df2_read['link'] = df2_read['link'].str.replace("https://", " ", regex = False)
    df2_read['link'] = df2_read['link'].str.replace("www.", " ", regex = False)
    df2_read['link'].str.strip()


    #2중 for문 돌려서 우리가 가지고있는 주소가 포함된 내용이 있는 항목이 발생할 때의 인덱스 값을 1로 변경해서 저장(Boolean으로 변경해도 상관 없음)
    arr1 = ["NS" for i in range(50)]
    idx =0
    global heat_count

    Os_Num = 0

    for x in df3_read['URL']:
        case = df2_read['link'].str.contains(x)
        # 포함되어 있는게 맞는 경우
        for y in case:
            if y == True:
                # OS 즉 우리의 사이트입니다.
                arr1[idx] = "OS"
                Os_Num = Os_Num + 1
            idx = idx + 1
        idx = 0

    #우리 사이트가 맞는 경우에는 OS체크.
    heat_count = arr1.count("OS")

    arr1 = pd.DataFrame(arr1, columns=['Status'])

    df2_read['Status'] = arr1
    df2_read['Result'] = heat_count
    df2_read['link']=df2_read['link2']
    del df2_read['link2']

    df_to_excel_each_sheet(df2_read, kw, f_name)
    wb.save('target_list.xlsx')
    print('완료!!')

for x in range(3):
    print(1);

# get_occu_rate(organic_results_df1, 'World Mission Society Church of God', '81_영어_[맨섬].xlsx', 83)
# get_occu_rate(organic_results_df1, 'World Mission Society Church of God', '104_영어_[저지섬].xlsx',104)
# get_occu_rate(organic_results_df1, 'World Mission Society Church of God','72_영어_[건지].xlsx', 74)


# def switch(key):
    #     telecom = {"011": "SKT", "016": "KT", "019": "LGU"}.get(key, "알 수 없는")
    #     print(f'당신은 {telecom} 사용자입니다.')
    #
    # firstNumber = input().split("-")[0]
    # switch(firstNumber)

# for kw in kws:
#     params.update(q=kw)
#
# # 파라메터 값에 따라 구글 검색결과 results로 반환하기
# search = GoogleSearch(params)
# results = search.get_dict()
#
# # 검색결과에서 자연 검색결과 추출
# organic_results = results['organic_results']
# # json을 Dataframe으로 변환
# organic_results_df = json_to_df(organic_results)
# # df를 엑셀로 추출
# df_to_excel_each_sheet(organic_results_df, kw, "google_serp/serp.xlsx")


# #카테고리화 하기위한 일들 'twitter', 'wiki', 'youtube','facebook','watv' 'vimeo'
# choicelist = ['twitter', 'wiki', 'youtube','facebook','watv'] 굳이 이 방식 말고 위처럼 하면 될듯.
#
# conditionlist = [
#     (df2_read['link'].str.contains(choicelist[0])),
#     (df2_read['link'].str.contains(choicelist[1])),
#     (df2_read['link'].str.contains(choicelist[2])),
#     (df2_read['link'].str.contains(choicelist[3])),
#     (df2_read['link'].str.contains(choicelist[4])),
#     )
# ]
#
# arr2 = np.select(conditionlist, choicelist, default = 'NS')
# arr2 = pd.DataFrame(arr2)
#
# # 리스트 To 데이터프레임화하고 저장
# arr1 = pd.DataFrame(arr1, columns=['Status'])
# df2_read['Status'] = arr1
# df2_read['Category'] = arr2
#
# #체크가 필요한 리스트를 따로 지정
# df2_read['NeedToCheck'] = np.where((df2_read['Status'] == 'NS') & (df2_read['Category'] == 'NS'),"Yes", "No")
#
# #출처가 위키면 상태를 DS로 변경
# df2_read.loc[df2_read['Category']=='wiki', 'Status'] = 'DS'

# for a in range(10):
# print(a)
# if x:
# print(keycount)
# print(df_checklist.iloc[keycount,4])
# params 키 q의 값(키워드) 업데이트 하기
# params.update(q=kw)
# params.update(google_domain= x, gl,hl, key, api_key)
# keycount = keycount+1


