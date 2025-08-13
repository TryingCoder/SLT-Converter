def validate_file_format(file_path):
    """Validate if the file has a .qpw extension."""
    return file_path.lower().endswith('.qpw')

def log_message(message):
    """Log a message to the console."""
    print(message)

def batch_process_files(file_paths):
    """Process a list of file paths for conversion."""
    valid_files = [file for file in file_paths if validate_file_format(file)]
    if not valid_files:
        log_message("No valid .qpw files found for conversion.")
        return
    log_message(f"Found {len(valid_files)} valid .qpw files for conversion.")
    return valid_files