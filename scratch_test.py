import os
import sys
import datetime

# Add the app directory to path so we can import app.py
sys.path.append(r'C:\Users\abbkp\.gemini\antigravity\scratch\gatepass-app')

from app import init_excel, get_vehicles, add_vehicle, delete_vehicle, EXCEL_FILE
from app import get_users, add_user, update_user_status, delete_user_from_excel, hash_password

def run_tests():
    print("--- Starting Excel DB & Auth Tests ---")
    
    # 1. Clean up old excel if exists for clean test
    if os.path.exists(EXCEL_FILE):
        try:
            os.remove(EXCEL_FILE)
            print("Removed existing test Excel file.")
        except Exception as e:
            print(f"Notice: Could not remove excel: {e}")
        
    # 2. Init Excel (sets up Vehicles and Users sheets, seeds admin)
    init_excel()
    assert os.path.exists(EXCEL_FILE), "Excel file should be created!"
    print("Test 1: Excel Init successful.")
    
    # 3. Read seeded admin user
    users = get_users()
    assert len(users) == 1, f"Expected 1 seeded user, got {len(users)}"
    admin_user = users[0]
    assert admin_user["username"] == "admin"
    assert admin_user["role"] == "admin"
    assert admin_user["status"] == "approved"
    assert admin_user["passwordHash"] == hash_password("admin")
    print("Test 2: Seeded admin read successful.")
    
    # 4. Add new user with pending status
    new_user = {
        "username": "testuser",
        "passwordHash": hash_password("mypassword"),
        "role": "user",
        "status": "pending",
        "createdAt": datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    add_user(new_user)
    print("Test 3: New user added successfully.")
    
    # 5. Read back user list
    users = get_users()
    assert len(users) == 2, f"Expected 2 registered users, got {len(users)}"
    user_record = next((u for u in users if u["username"] == "testuser"), None)
    assert user_record is not None, "testuser record should exist"
    assert user_record["status"] == "pending"
    print("Test 4: User list read back successful.")
    
    # 6. Approve user
    update_success = update_user_status("testuser", "approved")
    assert update_success is True, "Update status should return True"
    
    # Verify status changed
    users = get_users()
    user_record = next((u for u in users if u["username"] == "testuser"), None)
    assert user_record["status"] == "approved", f"Expected approved status, got {user_record['status']}"
    print("Test 5: User approval update successful.")
    
    # 7. Delete user
    delete_success = delete_user_from_excel("testuser")
    assert delete_success is True, "Delete user should return True"
    
    # Verify user list is back to admin only
    users = get_users()
    assert len(users) == 1, f"Expected 1 user after deletion, got {len(users)}"
    assert users[0]["username"] == "admin"
    print("Test 6: User deletion successful.")
    
    # 8. Add vehicle with simplified schema (checking that vehicles list still works alongside users)
    mock_vehicle = {
        "vehicleNo": "KA-19-M-1111",
        "driverName": "Auth Driver",
        "punchNumber": "P-9999",
        "gatepassExpiry": "2026-07-25",
        "insuranceExpiry": "2026-07-30",
        "pucExpiry": "2026-07-05",
        "fitnessExpiry": "2026-08-15",
        "taxExpiry": "2026-07-01",
        "permitExpiry": "2026-09-20",
        "remark": "Test remark text",
        "dlPhoto": "mock_dl.png",
        "punchCard": "mock_pc.png",
        "createdAt": datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    add_vehicle(mock_vehicle)
    vehicles = get_vehicles()
    print(f"DEBUG: Vehicles found in Excel: {vehicles}")
    assert len(vehicles) == 1, f"Expected 1 vehicle, got {len(vehicles)}"
    assert vehicles[0]["vehicleNo"] == "KA-19-M-1111"
    print("Test 7: Vehicles database integration verified.")
    
    # Clean up vehicle
    delete_vehicle("KA-19-M-1111")
    
    print("\n--- ALL TESTS PASSED SUCCESSFULLY (AUTH + EXCEL LAYOUT INTEGRATED)! ---")

if __name__ == "__main__":
    run_tests()
