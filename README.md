# Functionalities

- Anonymization of data in Office files ***(docx, xlsx, pptx) using AES encryption;***
- Deanonymization of anonymized data, recovering the original data;
- File selection through dialog box;
- Informative messages about the status of the operation.

# Prerequisites

- Microsoft Office installed;
- .NET Framework 4.6 or higher;

# Installation

- Clone the GitHub repository to your local computer;
- Compile the project using Visual Studio;
- Upload the Microsoft Office add-in to the ***Office AddIns folder.***

# Use

- Open an Office document (docx, xlsx, pptx);
- On the Office ribbon, access the ***Anonymized Data tab;***
- To anonymize data, click on the ***Anonymize Data button;***
- Select the file to be anonymized;
- Enter the secret password to encrypt data;
- Click OK to start anonymization;
- The original file will be preserved and a new encrypted file with the name ***Anonymized_ + file_name will be created;***
- To de-anonymize data, click the De-anonymize Data button;
- Select the anonymized file;
- Enter the secret password used to encrypt the data;
- Click OK to start deanonymization;
- The original file will be restored from the anonymized file.

# Security

- The tool uses AES encryption with a 256-bit key to guarantee the security of anonymized data.
- The secret password used to encrypt data is crucial to data protection and should not be shared with anyone.

# Comments

- The tool does not modify the original file, a new encrypted file is created and the original file is preserved.
- It is important to remember that data anonymization can make data unrecoverable if the secret password is lost.
- The tool is provided as is and does not provide any warranty.

# Contributions

- We value the community's opinion and contribution, suggestions, bug fixes and new features can be submitted through GitHub.
