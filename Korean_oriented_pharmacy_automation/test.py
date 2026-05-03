import os
os.environ["FLAGS_use_mkldnn"] = "0"

from paddleocr import PaddleOCR

# OCR 객체 생성 (PaddleOCR 2.x API)
ocr = PaddleOCR(use_angle_cls=True, lang='korean')

# 이미지 경로 (한글 경로 지원)
image_path = r"C:\Users\COM\Downloads\테스트사진.png"

# OCR 실행
result = ocr.ocr(image_path, cls=True)

# 결과 출력
for line in result:
    for word_info in line:
        text = word_info[1][0]
        score = word_info[1][1]
        print(f"{text} (신뢰도: {score:.2f})")