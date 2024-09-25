import os

def read_file(filename):
    try:
        with open(filename, 'r') as file:
            return file.read()
    except FileNotFoundError:
        print("File not found.")
        return None

def get_user_input():
    return input("Enter a command: ")

def execute_command(command):
    os.system(command)

def main():
    filename = get_user_input()
    file_content = read_file(filename)
    
    if file_content:
        print("File content:")
        print(file_content)
    
    command = get_user_input()
    execute_command(command)

if __name__ == "__main__":
    main()

class MyClass(object):
    def __init__(self):
        self.message = 'Hello'
        return self  # Noncompliant
