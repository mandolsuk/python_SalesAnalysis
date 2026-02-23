# -*- coding: utf-8 -*-
"""
재판매 마진 분석 스크립트
Resale Margin Analysis Tool

평택·천안 공실 오피스 활용 중계/재판매 사업을 위한
품목별 수익성 분석 도구

Usage:
    python resale_margin_analysis.py
"""

import pandas as pd


def calculate_margin(purchase_price, selling_price, platform_fee_rate=0.06,
                     shipping_cost=3000, packaging_cost=1000):
    """품목별 마진을 계산합니다.

    Args:
        purchase_price: 매입가 (원)
        selling_price: 판매가 (원)
        platform_fee_rate: 플랫폼 수수료율 (기본 6%)
        shipping_cost: 배송비 (원)
        packaging_cost: 포장비 (원)

    Returns:
        dict: 마진 분석 결과
    """
    platform_fee = int(selling_price * platform_fee_rate)
    total_cost = purchase_price + platform_fee + shipping_cost + packaging_cost
    profit = selling_price - total_cost
    margin_rate = (profit / selling_price * 100) if selling_price > 0 else 0

    return {
        "매입가": purchase_price,
        "판매가": selling_price,
        "플랫폼수수료": platform_fee,
        "배송비": shipping_cost,
        "포장비": packaging_cost,
        "총비용": total_cost,
        "순이익": profit,
        "마진율(%)": round(margin_rate, 1),
    }


def analyze_monthly_revenue(items, office_rent=400000):
    """월 예상 수익을 분석합니다.

    Args:
        items: 품목별 거래 정보 리스트
            각 항목은 dict로 {품목, 월거래량, 평균매입가, 평균판매가} 포함
        office_rent: 공실 오피스 월 임대료 (원)

    Returns:
        pandas.DataFrame: 월 수익 분석 결과
    """
    results = []
    for item in items:
        margin = calculate_margin(item["평균매입가"], item["평균판매가"])
        monthly_profit = margin["순이익"] * item["월거래량"]
        monthly_revenue = item["평균판매가"] * item["월거래량"]
        results.append({
            "품목": item["품목"],
            "월거래량": item["월거래량"],
            "평균매입가": item["평균매입가"],
            "평균판매가": item["평균판매가"],
            "건당순이익": margin["순이익"],
            "마진율(%)": margin["마진율(%)"],
            "월매출": monthly_revenue,
            "월순이익": monthly_profit,
        })

    df = pd.DataFrame(results)

    total_monthly_profit = df["월순이익"].sum()
    net_after_rent = total_monthly_profit - office_rent

    return df, total_monthly_profit, net_after_rent


def main():
    """메인 실행 함수: 샘플 데이터로 마진 분석을 실행합니다."""

    print("=" * 60)
    print("  재판매 마진 분석 도구")
    print("  평택·천안 공실 오피스 활용 사업 시뮬레이션")
    print("=" * 60)

    # 샘플 품목 데이터
    sample_items = [
        {"품목": "중고 스마트폰", "월거래량": 15, "평균매입가": 200000, "평균판매가": 300000},
        {"품목": "중고 노트북", "월거래량": 5, "평균매입가": 300000, "평균판매가": 450000},
        {"품목": "브랜드 스니커즈", "월거래량": 10, "평균매입가": 80000, "평균판매가": 150000},
        {"품목": "게임기(닌텐도/PS)", "월거래량": 8, "평균매입가": 150000, "평균판매가": 220000},
        {"품목": "사무용 가구", "월거래량": 5, "평균매입가": 30000, "평균판매가": 100000},
        {"품목": "유아용품", "월거래량": 10, "평균매입가": 20000, "평균판매가": 50000},
    ]

    # 단품 마진 분석
    print("\n[1] 품목별 단품 마진 분석")
    print("-" * 60)
    for item in sample_items:
        margin = calculate_margin(item["평균매입가"], item["평균판매가"])
        print(f"\n▶ {item['품목']}")
        for key, value in margin.items():
            if key == "마진율(%)":
                print(f"  {key}: {value}%")
            else:
                print(f"  {key}: {value:,}원")

    # 월 수익 분석
    office_rent = 400000  # 공실 오피스 월 임대료
    df, total_profit, net_profit = analyze_monthly_revenue(
        sample_items, office_rent
    )

    print("\n\n[2] 월 수익 시뮬레이션")
    print("-" * 60)
    print(df.to_string(index=False))
    print("-" * 60)
    print(f"월 총 순이익: {total_profit:,}원")
    print(f"오피스 임대료: {office_rent:,}원")
    print(f"임대료 차감 후 순이익: {net_profit:,}원")

    # 엑셀 출력
    output_file = "resale_analysis_output.xlsx"
    df.to_excel(output_file, index=False)
    print(f"\n분석 결과가 '{output_file}'에 저장되었습니다.")


if __name__ == "__main__":
    main()
