This was repository for paper "Development of an Engineering Drawing Detection and Extraction Algorithm for Quality Inspection Using Deep Neural Networks"

There are two pretained models in this repository:
- yolov8_view.pt : to detect views in engineering drawing
- yolov8_annotation.pt : to detect dimensions in engineering drawing view

The engineering drawing dimension extraction flow and libraries used:
1. View detection: YOLO-v8
2. Dimension detection: YOLO-v8
3. Text recognition: OpenCV and Pytesseract
4. Information block processing: Camelot
5. Output generation: OpenPyXL

This study is supported by the Department of Industrial Engineering, Faculty of Industrial Technology, Bandung Institute of Technology (ITB) and CV Cipta Sinergi Manufacturing (CV CSM). 
