curl -X POST "http://127.0.0.1:8100/api/report/basic-info" \
  -H "Content-Type: application/json" \
  -d '{
        "project_name": "智慧数据中心巡检项目",
        "room_name": "A区主机房",
        "year": 2025,
        "quarter": "4季度",
        "report_date": "2025-10-20",
        "report_person": "张三"
      }'