🧠 Facebook Post Engagement Metrics Extractor
📋 Overview

This Python project automates the extraction of engagement metrics — likes, comments, and shares — from Facebook posts using Selenium.
It supports both single and bulk post URLs and features automatic login handling using either saved cookies or Facebook credentials from a .env file.

The script saves all extracted data into a structured Excel file (Facebook_Posts.xlsx), making it ideal for research, analytics, or content performance tracking.

-------------------------------------------------------------------------------------------------------------

🚀 Key Features

🔐 Automatic Facebook Login

Loads stored cookies for seamless login.

If cookies are missing, expired, or invalid, the script automatically logs in using credentials from the .env file and saves fresh cookies.

📄 Bulk URL Processing

Reads post URLs from a text file (facebook_posts.txt).

Supports both single and multiple Facebook post links.

📊 Engagement Data Extraction

Extracts Likes, Comments, and Shares from each post.

Works with both personal and public post links.

🧠 Human-like Interaction

Mimics real user behavior by introducing random delays and typing patterns to reduce detection by Facebook’s automation systems.

📁 Excel Export

Saves collected results to Facebook_Posts.xlsx with clean column formatting.

-------------------------------------------------------------------------

🧰 Requirements

Make sure you have the following installed:

Python 3.8+

Google Chrome browser

ChromeDriver (matching your Chrome version)

------------------------------------------------------------------------------------------------

Facebook account credentials

⚙️ Installation
1. Clone the repository
git clone https://github.com/Avinashricky211/Extract-Facebook-Posts.git
cd Extract-Facebook-Posts

2. Install dependencies
pip install -r requirements.txt

3. Create a .env file

In the root folder of your project, create a .env file and add your Facebook credentials:

FACEBOOK_EMAIL=your_facebook_email

FACEBOOK_PASSWORD=your_facebook_password



📄 Usage
1. Add post URLs

Create a text file named facebook_posts.txt in the same directory as your script.
Each line should contain one Facebook post URL:

https://www.facebook.com/post1
https://www.facebook.com/post2
https://www.facebook.com/post3

2. Run the script

Execute the main file:

python main.py

3. What happens when you run it

The script tries to log in using saved cookies (facebook_cookies.pkl).

If cookies are missing or expired, it automatically logs in with your .env credentials and saves new cookies for next time.

The script loads each URL from facebook_posts.txt.

It extracts the Like, Comment, and Share counts from each post.

All results are saved to an Excel file named Facebook_Posts.xlsx.

📊 Output Example
LIKES	COMMENTS	SHARES	URL	CREATEDATE
125	43	18	https://facebook.com/post1
	24/10/25 09:30 AM
97	22	11	https://facebook.com/post2
	24/10/25 09:37 AM

📁 Output File: Facebook_Posts.xlsx
-----------------------------------------------------------------------------------------------------
🧩 Project Structure
Extract-Facebook-Posts/
│
├── main.py                 # Main script to execute the scraping
├── facebook_posts.txt      # File containing post URLs to scrape
├── Facebook_Posts.xlsx     # Output Excel file (auto-generated)
├── facebook_cookies.pkl    # Saved login session cookies (auto-generated)
├── .env                    # Your Facebook credentials (created manually)
├── requirements.txt        # List of required Python packages
└── README.md               # Project documentation

⚙️ Main Components

FacebookDataCollector → Handles login, cookies, scraping, and automation

extract_metrics_with_xpath() → Extracts engagement metrics using custom XPath + JavaScript logic

save_results_to_excel() → Writes final output to Excel with proper formatting

facebook_posts.txt → Input file for URLs to scrape

.env → Stores Facebook login credentials securely

--------------------------------------------------------------------

💡 Tips

Ensure you’re using the latest version of Chrome and ChromeDriver.

For large batches of URLs, use reasonable delays to avoid being flagged by Facebook.

Run the script in non-headless mode (default) to debug login or scraping issues.

---------------------------------------------------------------------------------------------------

⚠️ Disclaimer

This project is intended for educational and research purposes only.
Automating interactions with Facebook may violate its Terms of Service.
Use responsibly and only on accounts or pages you own or have permission to access.

--------------------------------------------------------------------------------------------------------

✨ Author & Contribution Notes

Hi there! 👋
I’m Avinash Ricky, the creator of this project.

I’m passionate about automation, AI, and machine learning, and I enjoy building practical tools that make data collection and analysis more efficient.
This Facebook engagement extractor is one of many projects I’ve developed to simplify real-world workflows through intelligent automation.

💡 Contributions

Contributions, suggestions, and feature improvements are always welcome!

You can fork this repository, make your changes, and open a pull request.

Whether it’s optimizing the scraping logic, improving data processing, or extending it with AI insights — your ideas are appreciated.

💼 Collaboration & Opportunities

I’m open to:

🚀 Freelance projects

💻 Full-time or contract opportunities

🤝 Collaborative research or open-source contributions

If you’re looking for a developer who can handle Python, Automation, Web Scraping, Data Science, AI, or Machine Learning projects, feel free to reach out.

📫 Contact:

GitHub: https://github.com/Avinashricky211

LinkedIn: https://www.linkedin.com/in/yadlapalli-avinash-ricky/

Email: ricky48avinash@gmail.com

Let’s build something amazing together! 🚀
