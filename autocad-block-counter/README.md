# AutoCAD Block Counter & BOM Exporter

A professional C# plug-in for AutoCAD designed to automate block counting and data extraction. This tool solves the common engineering problem of manual inventory, significantly reducing the time required to create Bills of Materials (BOM).

## 🚀 Key Features

*   **Dynamic Block Support**: Correcty identifies the parent name of dynamic and anonymous blocks, avoiding cryptic system names (e.g., `*U...`).
*   **Selective Selection**: Allows users to count objects from any selection set (window, crossing, or individual selection).
*   **Transaction Safety**: Implements AutoCAD's `TransactionManager` to ensure drawing database stability and prevent crashes during object access.
*   **Excel Integration**: Generates a `.csv` file directly on the user's Desktop with UTF-8 encoding, ready for immediate use in Microsoft Excel.
## 🛠 Technical Stack

*   **Language**: C# (.NET)
*   **APIs**: AutoCAD .NET API (`accoremgd`, `acdbmgd`, `acmgd`)
*   **Output Format**: CSV (Semicolon delimited)

## 📖 How to Use

### 1. Build the Project
Open the `.csproj` file in Visual Studio and **Build** the solution to generate the `SelectCount.dll` library.

### 2. Load into AutoCAD
Inside the AutoCAD command line, type:
```text
NETLOAD
```
Browse and select the compiled SelectCount.dll from your disk.
### 3. Run the Command

Type the custom command:
```text
BlockCounter
```
### 4. Select the blocks on your screen. The program will count all unique block references and save a BOM.csv file on your Desktop.
📂 Project Structure

    BlockCounter.cs – Main logic for block counting and AutoCAD API interaction.

    SelectCount.csproj – Project configuration and Autodesk library dependencies.

    README.md – Project documentation.

Developed as part of my Pet Projects collection.
