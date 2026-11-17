#!/usr/bin/env python3
"""매일 아침 날씨를 MS Teams 채널에 올리는 스크립트.

- 날씨: Open-Meteo (API 키 불필요)
- 전송: Teams Workflows(Power Automate) 웹훅으로 Adaptive Card 전송
- 의존성: 표준 라이브러리만 사용 (Python 3.9+)

환경 변수
  TEAMS_WEBHOOK_URL  (필수) Teams Workflows 웹훅 URL
  WEATHER_LAT        (선택) 위도, 기본 37.5665 (서울)
  WEATHER_LON        (선택) 경도, 기본 126.9780
  WEATHER_LOCATION   (선택) 카드에 표시할 지역명, 기본 "서울"

--dry-run 옵션을 주면 Teams로 보내지 않고 카드 JSON만 출력한다.
"""
import json
import os
import sys
import urllib.parse
import urllib.request
from datetime import datetime
from zoneinfo import ZoneInfo

KST = ZoneInfo("Asia/Seoul")
WEEKDAYS = "월화수목금토일"

# WMO weather code -> (이모지, 설명)
WMO = {
    0: ("☀️", "맑음"),
    1: ("🌤️", "대체로 맑음"),
    2: ("⛅", "구름 조금"),
    3: ("☁️", "흐림"),
    45: ("🌫️", "안개"),
    48: ("🌫️", "서리 안개"),
    51: ("🌦️", "약한 이슬비"),
    53: ("🌦️", "이슬비"),
    55: ("🌦️", "강한 이슬비"),
    56: ("🌧️", "어는 이슬비"),
    57: ("🌧️", "강한 어는 이슬비"),
    61: ("🌧️", "약한 비"),
    63: ("🌧️", "비"),
    65: ("🌧️", "강한 비"),
    66: ("🌧️", "어는 비"),
    67: ("🌧️", "강한 어는 비"),
    71: ("🌨️", "약한 눈"),
    73: ("🌨️", "눈"),
    75: ("❄️", "강한 눈"),
    77: ("🌨️", "싸락눈"),
    80: ("🌦️", "약한 소나기"),
    81: ("🌦️", "소나기"),
    82: ("⛈️", "강한 소나기"),
    85: ("🌨️", "약한 소낙눈"),
    86: ("❄️", "강한 소낙눈"),
    95: ("⛈️", "뇌우"),
    96: ("⛈️", "우박 동반 뇌우"),
    99: ("⛈️", "강한 우박 동반 뇌우"),
}


def fetch_weather(lat: float, lon: float) -> dict:
    params = urllib.parse.urlencode(
        {
            "latitude": lat,
            "longitude": lon,
            "current": "temperature_2m,apparent_temperature,relative_humidity_2m,weather_code,wind_speed_10m",
            "daily": "weather_code,temperature_2m_max,temperature_2m_min,precipitation_probability_max,precipitation_sum",
            "timezone": "Asia/Seoul",
            "wind_speed_unit": "ms",
            "forecast_days": 1,
        }
    )
    url = f"https://api.open-meteo.com/v1/forecast?{params}"
    with urllib.request.urlopen(url, timeout=15) as resp:
        return json.load(resp)


def advice(daily_code: int, rain_prob: int, t_min: float, t_max: float) -> str:
    tips = []
    snowy = daily_code in (71, 73, 75, 77, 85, 86)
    if snowy:
        tips.append("눈 예보가 있어요. 출근길 미끄럼 조심하세요.")
    elif rain_prob >= 60:
        tips.append("우산 챙기세요.")
    elif rain_prob >= 30:
        tips.append("비 올 수 있으니 작은 우산 하나 챙기면 좋아요.")
    if t_min <= 0:
        tips.append("아침 기온이 영하예요. 따뜻하게 입으세요.")
    elif t_max >= 30:
        tips.append("낮에 많이 더워요. 물 자주 드세요.")
    if t_max - t_min >= 10:
        tips.append("일교차가 커서 겉옷이 필요해요.")
    return " ".join(tips) or "특별히 챙길 건 없어요."


def build_card(weather: dict, location: str, now: datetime) -> dict:
    cur, day = weather["current"], weather["daily"]
    daily_code = day["weather_code"][0]
    t_max = day["temperature_2m_max"][0]
    t_min = day["temperature_2m_min"][0]
    rain_prob = day["precipitation_probability_max"][0] or 0
    rain_sum = day["precipitation_sum"][0] or 0

    cur_emoji, cur_desc = WMO.get(cur["weather_code"], ("🌡️", "정보 없음"))
    day_emoji, day_desc = WMO.get(daily_code, ("🌡️", "정보 없음"))
    date_str = f"{now.month}월 {now.day}일 ({WEEKDAYS[now.weekday()]})"

    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "msteams": {"width": "Full"},
        "body": [
            {
                "type": "TextBlock",
                "text": f"{day_emoji} {date_str} {location} 날씨",
                "weight": "Bolder",
                "size": "Large",
                "wrap": True,
            },
            {
                "type": "TextBlock",
                "text": f"지금 {cur['temperature_2m']:.0f}°C · {cur_emoji} {cur_desc} "
                f"(체감 {cur['apparent_temperature']:.0f}°C)",
                "size": "Medium",
                "spacing": "Small",
                "wrap": True,
            },
            {
                "type": "FactSet",
                "facts": [
                    {"title": "오늘 날씨", "value": day_desc},
                    {"title": "최저 / 최고", "value": f"{t_min:.0f}°C / {t_max:.0f}°C"},
                    {"title": "강수확률", "value": f"{rain_prob}%"
                     + (f" (예상 {rain_sum:.1f}mm)" if rain_sum > 0 else "")},
                    {"title": "습도", "value": f"{cur['relative_humidity_2m']}%"},
                    {"title": "바람", "value": f"{cur['wind_speed_10m']:.1f} m/s"},
                ],
            },
            {
                "type": "TextBlock",
                "text": advice(daily_code, rain_prob, t_min, t_max),
                "wrap": True,
                "spacing": "Medium",
            },
        ],
    }
    # Teams Workflows 웹훅이 받는 형식
    return {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "contentUrl": None,
                "content": card,
            }
        ],
    }


def send_to_teams(webhook_url: str, payload: dict) -> None:
    req = urllib.request.Request(
        webhook_url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        if resp.status not in (200, 202):
            raise RuntimeError(f"Teams 전송 실패: HTTP {resp.status}")


def main() -> int:
    dry_run = "--dry-run" in sys.argv
    lat = float(os.environ.get("WEATHER_LAT", "37.5665"))
    lon = float(os.environ.get("WEATHER_LON", "126.9780"))
    location = os.environ.get("WEATHER_LOCATION", "서울")

    payload = build_card(fetch_weather(lat, lon), location, datetime.now(KST))

    if dry_run:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0

    webhook_url = os.environ.get("TEAMS_WEBHOOK_URL")
    if not webhook_url:
        print("TEAMS_WEBHOOK_URL 환경 변수가 없습니다.", file=sys.stderr)
        return 1
    send_to_teams(webhook_url, payload)
    print("전송 완료")
    return 0


if __name__ == "__main__":
    sys.exit(main())
