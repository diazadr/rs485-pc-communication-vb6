# RS485 PC-to-PC Communication using Visual Basic

![Project Status](https://img.shields.io/badge/status-completed-brightgreen) [![License](https://img.shields.io/badge/license-MIT-blue)](./LICENSE)

This tutorial demonstrates how to create a chat application in Visual Basic that enables communication between PCs using RS485 as the transmission medium.

## Technologies Used
- **RS485 USB Adapter**: Hardware for serial communication.
- **Visual Basic**: Used to build the chat application interface.

## Features
- **PC-to-PC Communication**: Enables message exchange between multiple computers via RS485.
- **Serial Communication Handling**: Uses MSComm for data transmission and reception.
- **Message Parsing**: Processes and validates messages with specific protocol formatting.

## Demo

### **System Overview**
RS485 allows multiple devices to communicate over a single bus. The chat application is designed to send and receive messages based on predefined addressing and formatting rules.

### **Message Formatting**
Messages are structured as follows:
```
*SourceAddress|DestinationAddress|Message#
```
- `*` : Start of the message
- `SourceAddress` : Address of the sender
- `DestinationAddress` : Address of the receiver ("A" for broadcast to all)
- `Message` : The actual message
- `#` : End of the message

### **Parsing the Data in Visual Basic**
To extract relevant message information, the following parsing functions are used:
- **Len()**: Determines the length of the message.
- **Left()**: Extracts the first characters from the left.
- **Mid()**: Extracts a substring from a specific position.
- **Right()**: Extracts the last characters from the right.

### **Visual Basic Form Interface**
Below is an example of the Visual Basic user interface for the chat application:

<div align="left">
  <img src="https://github.com/user-attachments/assets/example-chat-ui.png" alt="Chat Application Interface" width="600">
</div>

### **Message Reception and Processing**
Messages are only processed if they meet the following conditions:
1. The message starts with `*` and ends with `#`.
2. The destination address matches the receiving PC’s address or is set to broadcast (`A`).
3. Parsed messages are displayed in the chat window with the sender’s address.

### **Message Transmission**
To send a message, the application:
1. Concatenates the required components into the correct format.
2. Sends the formatted string over the RS485 network.

### **Visual Basic Form Interface PC1**  
Below is an example of the Visual Basic user interface:

<div align="left">
  <img src="https://github.com/user-attachments/assets/75ea6f2a-a511-41b9-bbab-b56214974e2f" alt="Form Interface Screenshot" width="600">
</div>

### **Visual Basic Form Interface PC2**  
Below is an example of the Visual Basic user interface:

<div align="left">
  <img src="https://github.com/user-attachments/assets/8bfc0cd7-981a-4923-840f-c12fd18cb2ea" alt="Form Interface Screenshot" width="600">
</div>

## MSComm Setup in Visual Basic
To configure MSComm for serial communication:
1. Click **Project**, select **Components** (or press **Ctrl + T**).
2. Check **Microsoft Comm Control 6.0 (SP6)** and click **OK**.
3. Configure the `CommPort` and serial settings:
   - Baudrate = 9600
   - Data = 8 bits
   - Parity = None
   - Stop bit = 1

## Installation and Usage
1. **Connect RS485 adapters** to the PCs.
2. **Set up MSComm** in Visual Basic.
3. **Run the application** and start sending messages between PCs.

## Project Status
This project is **completed** and will not be further developed.

## Contributions
Feel free to submit an **issue** or create a **pull request** if you wish to contribute.

## License
This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for more details.

