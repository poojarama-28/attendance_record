import cv2
import os
import datetime
import pandas as pd

# Directory to save student images
IMAGE_DIR = "student_images"
os.makedirs(IMAGE_DIR, exist_ok=True)

# Initialize face detector
face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")

# Initialize attendance log Excel file
ATTENDANCE_FILE = "attendance_log.xlsx"
if not os.path.exists(ATTENDANCE_FILE):
    df = pd.DataFrame(columns=["Roll Number", "Name", "Timestamp", "Status"])
    df.to_excel(ATTENDANCE_FILE, index=False)

def capture_student_image(student_roll, student_name):
    cap = cv2.VideoCapture(0)
    while True:
        ret, frame = cap.read()
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = face_cascade.detectMultiScale(gray, 1.3, 5)
        
        for (x, y, w, h) in faces:
            # Draw a rectangle around the detected face
            cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 255, 0), 2)
            cv2.putText(frame, "Face Detected", (x, y-10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 255, 0), 2)
            
            # Save the face image
            image_path = os.path.join(IMAGE_DIR, f"{student_roll}_{student_name}.jpg")
            cv2.imwrite(image_path, frame[y:y+h, x:x+w])
            
            # Mark attendance
            mark_attendance(student_roll, student_name)
            print(f"Image saved for Student Roll: {student_roll}, Name: {student_name}")
            
            # Release the camera and close windows
            cap.release()
            cv2.destroyAllWindows()
            return
        
        # Display the live video feed
        cv2.imshow("Capture Image - Press 'Q' to Exit", frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    
    # Release the camera and close windows
    cap.release()
    cv2.destroyAllWindows()

# Attendance log function
def mark_attendance(student_roll, student_name):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    attendance_data = {
        "Roll Number": [student_roll],
        "Name": [student_name],
        "Timestamp": [timestamp],
        "Status": ["Present"]
    }
    
    # Save to Excel
    save_to_excel(attendance_data)
    print(f"Attendance marked for {student_name} (Roll: {student_roll}) at {timestamp}")

# Save attendance data to Excel
def save_to_excel(data):
    df = pd.DataFrame(data)
    if not os.path.exists(ATTENDANCE_FILE):
        df.to_excel(ATTENDANCE_FILE, index=False)
    else:
        existing_df = pd.read_excel(ATTENDANCE_FILE)
        updated_df = pd.concat([existing_df, df], ignore_index=True)
        updated_df.to_excel(ATTENDANCE_FILE, index=False)

# View attendance records
def view_attendance():
    if os.path.exists(ATTENDANCE_FILE):
        df = pd.read_excel(ATTENDANCE_FILE)
        print("\n--- Attendance Records ---")
        print(df)
    else:
        print("No attendance records found.")

# Interactive menu
def main_menu():
    while True:
        print("\n--- Smart Attendance System ---")
        print("1. Capture Student Image and Mark Attendance")
        print("2. View Attendance Records")
        print("3. Exit")
        choice = input("Enter your choice (1-3): ")
        
        if choice == "1":
            student_roll = input("Enter Student Roll Number: ")
            student_name = input("Enter Student Name: ")
            capture_student_image(student_roll, student_name)
        elif choice == "2":
            view_attendance()
        elif choice == "3":
            print("Exiting the program. Have a great day!")
            break
        else:
            print("Invalid choice. Please try again.")

# Example usage
if __name__ == "__main__":
    main_menu()