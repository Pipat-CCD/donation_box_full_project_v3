"""
สคริปต์ช่วยเตรียมข้อมูลสำหรับ Google Sheets
รันเพื่อแสดงชื่อคอลัมน์ที่ต้องสร้างในแต่ละชีต
"""
BOX_COLUMNS = ["box_id", "original_no", "province_zone", "subzone", "store_name", "floor", "contact", "landmark", "gps_lat", "gps_lng", "google_maps_url", "status", "cycle_months", "last_collect_month", "last_collect_day", "last_collect_date", "next_due_date", "note"]
COLLECTION_COLUMNS = ["record_id", "box_id", "store_name", "collect_date", "collector", "amount", "box_condition", "result_status", "photo_url", "gps_lat", "gps_lng", "note", "created_at"]
MISSING_COLUMNS = ["report_id", "box_id", "store_name", "found_date", "contact_date", "contact_person", "detail", "photo_url", "gps_lat", "gps_lng", "case_status", "created_at"]

print("boxes:")
print(",".join(BOX_COLUMNS))
print("\ncollections:")
print(",".join(COLLECTION_COLUMNS))
print("\nmissing_reports:")
print(",".join(MISSING_COLUMNS))
