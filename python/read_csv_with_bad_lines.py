import pandas as pd
import io

def read_csv_with_bad_lines(filepath, encoding='cp949', delimiter=','):
    """bad line을 skip하지 않으면서도 정상적으로 csv 파일을 읽어 드리는 방법
    - filepath : str : csv 파일을 읽어 올 파일 경로로 확장자를 포함함
    - encoding='cp949' : 기본 인코딩 스타일은 엑셀에서 csv를 저장하는 경우인 'cp949'
    - delimiter=';' : str : csv 파일을 읽어드리는 구분자

    - dtype=str로 모든 데이터를 문자열로 읽음

    - return dataframe을 리턴한다"""

    
    # 빈 데이터프레임 생성
    df = pd.DataFrame()

    with open(filepath, 'r', encoding=encoding) as file:
        lines = file.readlines()

    for line in lines:
        try:
            # 각 라인을 시도하여 데이터프레임에 추가, dtype=str로 모든 데이터를 문자열로 읽음
            temp_df = pd.read_csv(io.StringIO(line), header=None, dtype=str, delimiter=delimiter)
            df = pd.concat([df, temp_df], ignore_index=True)
        except pd.errors.ParserError:
            # 구문 오류가 발생하면 잘못된 라인을 데이터프레임에 추가
            line_data = [str(item) for item in line.split(delimiter)]  # 모든 항목을 문자열로 변환
            df = pd.concat([df, pd.DataFrame([line_data])], ignore_index=True)

    return df
