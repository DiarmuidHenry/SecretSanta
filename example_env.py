# Example of a functioning env.py file

import os

# Email functionality
os.environ["EMAIL_FUNCTIONALITY"] = "True"   

# Email server settings
os.environ["SMTP_SERVER"] = "smtp.gmail.com"
os.environ["SMTP_PORT"] = "XXX"

# Credentials
os.environ["EMAIL_SENDER"] = "emailaddress@email.com"
os.environ["EMAIL_PASSWORD"] = "emailpasswordhere"
 
# Print details settings      
os.environ["PRINT_CYCLES"] = "True"
os.environ["PRINT_NAMED_CYCLES"] = "False"