import pandas as pd
import matplotlib.pyplot as plt


def dataFrame_preprocessing():
    """데이터를 DataFrame 으로 읽어들인 후 결과값을 위한 데이터 전처리 작업 진행"""
    # ---------------------------------------------------------------------------------------
    """data load"""
    main_Data = pd.read_excel('데이터.xlsx')

    # ---------------------------------------------------------------------------------------
    """산업별 데이터 개수 지정 dict -> 추후 데이터 개수가 5개 미만인 데이터 제외 시 사용"""
    data_frequency_by_Industry = dict()
    for e, i in enumerate(main_Data['업종_중분류']):
        if i not in data_frequency_by_Industry:
            data_frequency_by_Industry[i] = 0
        else:
            data_frequency_by_Industry[i] += 1

    # ---------------------------------------------------------------------------------------
    """종목코드 삭제"""
    main_Data.drop(labels=['종목코드'], axis=1, inplace=True)

    # ---------------------------------------------------------------------------------------
    """업종별로 정렬"""
    main_Data.sort_values(by=['업종_중분류'], axis=0, ignore_index=True, inplace=True)

    # ---------------------------------------------------------------------------------------
    """산업별 중분류 DataFrame 생성"""
    classification_Data = pd.DataFrame(columns=['업종_중분류'])

    # ---------------------------------------------------------------------------------------
    """
    표본이 n개 이하인 산업은 제외
        ex) input = 10
                    10
                    10
                    11
                    11
                    ...
    """

    count = 0
    cut_data_num = 5
    for num in range(len(main_Data.index)):
        if data_frequency_by_Industry[main_Data['업종_중분류'][num]] >= cut_data_num:
            classification_Data.loc[count] = [main_Data['업종_중분류'][num]]
            count += 1
    print("데이터 개수가 5개 미만인 항목 제외")
    print("len(classification_Data) : {}\n".format(len(classification_Data)))
    print("classification_Data : \n {}".format(classification_Data))
    print()

    # ---------------------------------------------------------------------------------------
    """
    표본이 n개 이하인 산업은 제외
        ex) input = 10
                    11
                    13
                    ...
    """
    classification_Data = classification_Data.drop_duplicates(ignore_index=True)

    # ---------------------------------------------------------------------------------------
    """중분류별 업종, 업종명을 가지는 dict 생성"""
    # 업종, 업종명 dict
    category_csv = pd.read_excel('중분류.xlsx')

    industry_Dict = dict()
    for i, j in zip(category_csv['업종'], category_csv['업종명']):
        if i not in industry_Dict:
            industry_Dict[i] = j

    industry_class_Data = []
    for i in classification_Data['업종_중분류']:
        industry_class_Data.append(i)

    industry_class_Data = sorted(industry_class_Data)
    print("len(industry_class_Data)", len(industry_class_Data))
    print("industry_class_Data :", industry_class_Data)

    return main_Data, data_frequency_by_Industry, classification_Data, industry_Dict, industry_class_Data


def graph_sales(main_data, result_quartile):
    """매출로 분석한 결과를 그래프로 그리기"""
    box_plot_data = []
    box_plot_code = []
    for e, industry in enumerate(result_quartile):
        graph_data = pd.DataFrame(columns=['Industry Code', 'Sales'])
        count = 0
        for num in range(len(main_data)):
            if main_data['업종_중분류'][num] == int(industry[0]):
                graph_data.loc[count] = [main_data['업종_중분류'][num], main_data['3분기 누적 매출액 증감률'][num]]
                count += 1

        box_plot_code.append(industry[0])
        box_plot_data.append(graph_data['Sales'])

    plt.boxplot(box_plot_data)
    font1 = {
        'size': 25
    }
    plt.title('Sales by industry', fontdict=font1)
    plt.xticks([i + 1 for i in range(len(box_plot_code))], box_plot_code, fontsize=15)
    plt.yticks(fontsize=15)
    plt.grid(True)
    plt.xlabel('Industry Code', fontdict=font1)
    plt.ylabel('y', fontdict=font1)
    plt.ylim(-2, 5)
    plt.show()

    # ---------------------------------------------------------------------------------------
    """상위, 하위 5개 산업 선정"""
    n = 5
    bottom_n_industy = result_quartile[:n]
    top_n_industy = sorted(result_quartile[-n:], reverse=True, key=lambda x: x[1])
    print("사분위수 기준 산업별 매출")
    print("bottom_{}_industy".format(str(n)))
    for i in bottom_n_industy:
        print("업종 코드 : {}, 3사분위 값: {:.4f}, 업종 이름 : {}".format(i[0], i[1], industry_dict[i[0]]))
    print()

    print("top_{}_industy".format(str(n)))
    for i in top_n_industy:
        print("업종 코드 : {}, 3사분위 값: {:.4f}, 업종 이름 : {}".format(i[0], i[1], industry_dict[i[0]]))
    print()

    # ---------------------------------------------------------------------------------------
    """top 5 산업에 대해서 하나의 fig 로 그래프 그리기"""
    box_plot_data = []
    box_plot_code = []
    for e, industry in enumerate(top_n_industy):
        graph_data = pd.DataFrame(columns=['Industry Code', 'Sales'])
        count = 0
        for num in range(len(main_data)):
            if main_data['업종_중분류'][num] == int(industry[0]):
                graph_data.loc[count] = [main_data['업종_중분류'][num], main_data['3분기 누적 매출액 증감률'][num]]
                count += 1

        box_plot_code.append(industry[0])
        box_plot_data.append(graph_data['Sales'])

    plt.boxplot(box_plot_data)
    plt.title('Top 5 industries with increased sales', fontdict=font1)
    plt.xticks([i + 1 for i in range(len(box_plot_code))], box_plot_code, fontsize=15)
    plt.xlabel('Industry Code', fontdict=font1)
    plt.ylabel('y', fontdict=font1)
    plt.ylim(-1, 12)
    # plt.ylim(-2, 4)
    plt.show()

    # ---------------------------------------------------------------------------------------
    """bottom 5 산업에 대해서 하나의 fig 로 그래프 그리기"""
    box_plot_data = []
    box_plot_code = []
    for e, industry in enumerate(bottom_n_industy):
        graph_data = pd.DataFrame(columns=['Industry Code', 'Sales'])
        count = 0
        for num in range(len(main_data)):
            if main_data['업종_중분류'][num] == int(industry[0]):
                graph_data.loc[count] = [main_data['업종_중분류'][num], main_data['3분기 누적 매출액 증감률'][num]]
                count += 1

        box_plot_code.append(industry[0])
        box_plot_data.append(graph_data['Sales'])

    plt.boxplot(box_plot_data)
    plt.title('Top 5 industries with reduced sales', fontdict=font1)
    plt.xticks([i + 1 for i in range(len(box_plot_code))], box_plot_code, fontsize=15)
    plt.xlabel('Industry Code', fontdict=font1)
    plt.ylabel('y', fontdict=font1)
    plt.ylim(-1, 1)
    # plt.ylim(-3, 2)
    plt.show()


def graph_profit_loss(main_data, result_quartile):
    # ---------------------------------------------------------------------------------------
    """당기순이익으로 분석한 결과를 그래프로 그리기"""
    box_plot_data = []
    box_plot_code = []
    for e, industry in enumerate(result_quartile):
        graph_data = pd.DataFrame(columns=['Industry Code', 'Profit Loss'])
        count = 0
        for num in range(len(main_data)):
            if main_data['업종_중분류'][num] == int(industry[0]):
                graph_data.loc[count] = [main_data['업종_중분류'][num], main_data['3분기 누적 당기순이익 증감률'][num]]
                count += 1

        box_plot_code.append(industry[0])
        box_plot_data.append(graph_data['Profit Loss'])

    plt.boxplot(box_plot_data)
    font1 = {
        'size': 25
    }
    plt.title('Net profit by industry', fontdict=font1)
    plt.xticks([i + 1 for i in range(len(box_plot_code))], box_plot_code, fontsize=15)
    plt.yticks(fontsize=15)
    plt.grid(True)
    plt.xlabel('Industry Code', fontdict=font1)
    plt.ylabel('y', fontdict=font1)
    plt.ylim(-4, 4)
    plt.show()

    # ---------------------------------------------------------------------------------------
    """상위, 하위 5개 산업 선정"""
    n = 5
    bottom_n_industy = result_quartile[:n]
    top_n_industy = sorted(result_quartile[-n:], reverse=True, key=lambda x: x[1])
    print("사분위수 기준 산업별 당기순이익")
    print("bottom_{}_industy".format(str(n)))
    for i in bottom_n_industy:
        print("업종 코드 : {}, 3사분위 값: {:.4f}, 업종 이름 : {}".format(i[0], i[1], industry_dict[i[0]]))
    print()

    print("top_{}_industy".format(str(n)))
    for i in top_n_industy:
        print("업종 코드 : {}, 3사분위 값: {:.4f}, 업종 이름 : {}".format(i[0], i[1], industry_dict[i[0]]))
    print()

    # ---------------------------------------------------------------------------------------
    """top 5 산업에 대해서 하나의 fig 로 그래프 그리기"""
    box_plot_data = []
    box_plot_code = []
    for e, industry in enumerate(top_n_industy):
        graph_data = pd.DataFrame(columns=['Industry Code', 'Profit Loss'])
        count = 0
        for num in range(len(main_data)):
            if main_data['업종_중분류'][num] == int(industry[0]):
                graph_data.loc[count] = [main_data['업종_중분류'][num], main_data['3분기 누적 당기순이익 증감률'][num]]
                count += 1

        box_plot_code.append(industry[0])
        box_plot_data.append(graph_data['Profit Loss'])

    plt.boxplot(box_plot_data)
    plt.title('Top 5 Industries with Increased Net Profit', fontdict=font1)
    plt.xticks([i + 1 for i in range(len(box_plot_code))], box_plot_code, fontsize=15)
    plt.xlabel('Industry Code', fontdict=font1)
    plt.ylabel('y', fontdict=font1)
    plt.ylim(-15, 40)
    # plt.ylim(-2, 4)
    plt.show()

    # ---------------------------------------------------------------------------------------
    """bottom 5 산업에 대해서 하나의 fig 로 그래프 그리기"""
    box_plot_data = []
    box_plot_code = []
    for e, industry in enumerate(bottom_n_industy):
        graph_data = pd.DataFrame(columns=['Industry Code', 'Profit Loss'])
        count = 0
        for num in range(len(main_data)):
            if main_data['업종_중분류'][num] == int(industry[0]):
                graph_data.loc[count] = [main_data['업종_중분류'][num], main_data['3분기 누적 당기순이익 증감률'][num]]
                count += 1

        box_plot_code.append(industry[0])
        box_plot_data.append(graph_data['Profit Loss'])

    plt.boxplot(box_plot_data)
    plt.title('Top 5 Industries with Decrease in Net Profit', fontdict=font1)
    plt.xticks([i + 1 for i in range(len(box_plot_code))], box_plot_code, fontsize=15)
    plt.xlabel('Industry Code', fontdict=font1)
    plt.ylabel('y', fontdict=font1)
    plt.ylim(-65, 20)
    # plt.ylim(-3, 2)
    plt.show()


def analysis(main_data, industry_class_data, criteria_item=""):
    """전체 산업별 3사분위수를 구한 후 시각화"""

    if criteria_item == 'Sales':
        criteria_item = '3분기 누적 매출액 증감률'

    elif criteria_item == 'Profit Loss':
        criteria_item = '3분기 누적 당기순이익 증감률'

    result_quartile = []
    for e, industry in enumerate(industry_class_data):
        temp = []
        for num in range(len(main_data)):
            if main_data['업종_중분류'][num] == int(industry):
                temp.append(main_data[criteria_item][num])

        series = pd.Series(temp)
        Q3 = series.quantile(.75)
        result_quartile.append([int(industry), Q3])
    print("len(result quartile) : {}".format(len(result_quartile)))
    result_quartile = sorted(result_quartile, key=lambda x: x[1])
    print("result quartile : {}\n".format(result_quartile))

    if criteria_item == '3분기 누적 매출액 증감률':
        graph_sales(main_data, result_quartile)

    elif criteria_item == '3분기 누적 당기순이익 증감률':
        graph_profit_loss(main_data, result_quartile)


if __name__ == '__main__':
    main_data, data_frequency_by_industry, classification_data, industry_dict, industry_class_data = dataFrame_preprocessing()
    analysis(main_data, industry_class_data, 'Sales')
    analysis(main_data, industry_class_data, 'Profit Loss')
