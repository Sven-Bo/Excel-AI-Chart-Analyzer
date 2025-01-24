# Excel Image & Chart Analyzer with OpenAI API

This tool lets you analyze selected images, charts, or shapes in Excel using OpenAI. It’s easy to use, and you’ll get insights right in your workbook.



## Video Tutorial (Coming Soon!)
[![YouTube Video](https://img.youtube.com/vi/XXX/0.jpg)](https://youtu.be/XXX)



## Features
- Works with charts, images, and shapes in Excel.
- Stores your custom prompt and results for easy reuse.
- Sends data to OpenAI’s API and fetches answers.



## How to Set It Up

### 1. Add the VBA Code
1. Open Excel and press `ALT + F11` to open the VBA Editor.
2. Click on **File → Import File...**.
3. Import these two files into your project:
   - `mJsonConverter.bas`
   - `mAIObjectAnalyzer.bas`

4. Save your file as a **macro-enabled workbook** (`.xlsm`).



### 2. Get Your OpenAI API Key
1. Go to [https://platform.openai.com/api-keys](https://platform.openai.com/api-keys).
2. Copy your API key.
3. In the VBA Editor, open `mAIObjectAnalyzer.bas`.
4. Find this line:
   ```vba
   Const API_KEY As String = "YOUR_OPENAI_API_KEY"
   ```
   Replace `"YOUR_OPENAI_API_KEY"` with your real API key.

If you forget to set your API key, the macro will remind you with a message and provide the link.



### 3. Use Named Ranges (Optional)
You can create two named ranges in your workbook for better usability:

- **`PromptCell`**: A cell where you can save or edit the prompt.
- **`OutputCell`**: A cell where the analysis result will appear.

#### To Create a Named Range:
1. Select a cell in Excel.
2. Go to **Formulas → Define Name**.
3. Enter `PromptCell` or `OutputCell` as the name and click **OK**.

If you don’t define these ranges, the macro will use a default prompt and show the result in a message box.



### 4. Make the Macro Available Everywhere (Optional)
If you want this macro in all your Excel files:

1. Open your **Personal Macro Workbook**:
   - In Excel, go to **View → Macros → Record Macro**.
   - Choose to store the macro in **Personal Macro Workbook**, then stop recording.

2. Import the two files into the **Personal Macro Workbook** in the VBA Editor.
3. Save and close the VBA Editor. Now, the macro will work in any Excel file.



### 5. Add a Button or Toolbar Shortcut
#### Add to Quick Access Toolbar:
1. Go to **File → Options → Quick Access Toolbar**.
2. Select **Macros** from the dropdown.
3. Add the macro (`AnalyzeSelectedObjectWithOpenAI`) to the toolbar.

#### Add a Button:
1. Go to the **Developer Tab** and insert a button.
2. Right-click the button, select **Assign Macro**, and choose `AnalyzeSelectedObjectWithOpenAI`.



## How to Use
1. Select an **image**, **shape**, or **chart** in Excel.
2. Run the macro:
   - From the Quick Access Toolbar or
   - By clicking the button you set up.

3. Edit the prompt (if needed) in the InputBox and click **OK**.
4. Check the result in the named range `OutputCell` (if defined) or in a message box.



## 🤝 Connect with Me
- 📺 **YouTube:** [CodingIsFun](https://youtube.com/c/CodingIsFun)
- 🌐 **Website:** [PythonAndVBA](https://pythonandvba.com)
- 💬 **Discord:** [Join the Community](https://pythonandvba.com/discord)
- 💼 **LinkedIn:** [Sven Bosau](https://www.linkedin.com/in/sven-bosau/)
- 📸 **Instagram:** [sven_bosau](https://www.instagram.com/sven_bosau/)

## 💖 Support 
If my tutorials help you, please consider [buying me a coffee](https://pythonandvba.com/coffee-donation).  
[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://pythonandvba.com/coffee-donation)

## 📬 Feedback & Collaboration
If you have ideas, feedback, or want to collaborate, reach out at contact@pythonandvba.com.  
![Logo](https://www.pythonandvba.com/banner-img)