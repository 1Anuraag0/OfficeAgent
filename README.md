# Office Agent AI

Office Agent AI is an intelligent Microsoft Word Add-in that allows users to manipulate their documents using natural language commands. Built with Next.js, React, and the Microsoft Office JavaScript API (Office.js), this taskpane add-in leverages AI (powered by the Gemini API) to seamlessly interact with your Word documents.

## Features

- **Natural Language Document Manipulation**: Tell the agent what you want to do, and it executes the actions using Office.js.
- **Rich Document Interactions**:
  - **Read Document Content**: Extract and analyze the text within the document.
  - **Text Insertion**: Automatically generate and insert text.
  - **Intelligent Deletion**: Delete specific paragraphs or entire page ranges.
  - **Content Filtering**: Keep only paragraphs that contain specific keywords or names.
  - **Formatting & Highlighting**: Format specific styles (e.g., Heading1, Normal), change fonts, or highlight all text.
  - **Find & Replace**: Intelligently locate and replace text throughout the document.
- **Modern UI**: Built with React, TailwindCSS, and Framer Motion for a responsive and beautiful taskpane experience.
- **AI Model Integration**: Configured to use powerful generative AI to parse user intent and map it to Office.js actions.

## Tech Stack

- **Framework**: [Next.js 16](https://nextjs.org/) (React 19)
- **Styling**: [TailwindCSS v4](https://tailwindcss.com/)
- **Animations**: [Framer Motion](https://www.framer.com/motion/)
- **Microsoft Office API**: [Office.js](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)
- **AI Provider**: Google Gemini API (`@google/genai`)

## Prerequisites

- Node.js (v20+ recommended)
- Microsoft Word (Desktop or Word Online)
- A valid Gemini API Key (or other configured AI provider keys)

## Installation & Setup

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Generate Local SSL Certificates:**
   Microsoft Office Add-ins require an HTTPS connection to run locally.
   ```bash
   npx office-addin-dev-certs install
   ```
   *Make sure to copy the generated certificates (`localhost.pem` and `localhost-key.pem`) into a `certificates` folder in the root of the project if they aren't generated there automatically.*

3. **Configure Environment Variables:**
   Create a `.env.local` file in the root directory and add your necessary API keys:
   ```env
   GEMINI_API_KEY=your_gemini_api_key_here
   ```

4. **Start the Development Server:**
   ```bash
   npm run dev
   ```
   This will start the Next.js application with HTTPS enabled on `https://localhost:3000`.

## Sideloading the Add-in in Word

To test the add-in in Microsoft Word:

### Word on the Web (Office Online)
1. Open a document in Word Online.
2. Go to the **Insert** tab -> **Add-ins**.
3. Select **Upload My Add-in**.
4. Choose the `manifest.xml` file from the root of this project.
5. The "Office Agent AI" button will appear in your Home ribbon.

### Word Desktop (Windows/Mac)
1. Create a shared network folder and place your `manifest.xml` inside it.
2. Open Word -> **File** -> **Options** -> **Trust Center** -> **Trust Center Settings** -> **Trusted Add-in Catalogs**.
3. Add your shared folder path, check "Show in Menu", and click OK.
4. Restart Word.
5. Go to **Insert** -> **Get Add-ins** -> **Shared Folder** and select the Office Agent AI.

*(Alternatively, use the `office-addin-debugging` package to automatically sideload).*

## Project Structure

- `manifest.xml`: The Office Add-in configuration file required by Microsoft Word.
- `src/lib/office.ts`: Contains the core Office.js wrapper functions that interact safely with the Word document (inserting, deleting, formatting, finding/replacing).
- `package.json`: Contains the scripts and dependencies required for the project.

## Development Notes

- **Demo Mode:** If the app is run outside of Microsoft Word (e.g., in a standard browser), it will fallback to a "Demo Mode" safely, allowing you to test the UI without Office.js throwing errors.
- **Context Handling:** The `office.ts` library includes safe wrappers around `Word.run` to handle context losses seamlessly during hot-module replacement (HMR).

## License

This project is licensed under the MIT License.
