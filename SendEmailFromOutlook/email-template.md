# Sample Email Template

Below is a sample template that demonstrates the personalization capabilities of the sendemailfromoutlook script. There is a formatted version of this example in the file sample.docx.
```
Hi [GivenName],

Thank you for attending our recent workshop on emerging technologies! It was a pleasure to have you join us for such an engaging session. Your participation and insights made the workshop even more valuable for everyone involved.

I wanted to take a moment to share some exciting updates and opportunities that might interest you:

**New Resources Coming Soon**

I'm thrilled to announce the upcoming launch of our new website and knowledge hub in Q1 2025. This platform will serve as your go-to resource for technical deep dives, best practices, implementation strategies, and practical insights from real-world scenarios. Be among the first to access exclusive content and join our growing community!

**Technical Resources for Professionals**

To help you streamline your operations, we're regularly publishing new tools and scripts on our GitHub repository. These resources are specifically designed for busy professionals like you, addressing common challenges in enterprise environments. Follow our repository to stay updated with the latest automation solutions that can save you valuable time.

**Stay Connected with Industry Leaders**

As part of our ongoing work, we continue to collaborate closely with industry leaders and technology partners. If there are specific questions or challenges you'd like us to explore with these teams on your behalf, don't hesitate to let me know. It's a great way to get deeper insights or practical advice tailored to your operations!

**New Training Content Coming Soon**

We're hard at work developing new training materials that go even further into practical use cases, advanced features, and innovative integrations. Stay tuned—2025 is going to be an exciting year for technology professionals!

**Start the New Year with Special Offers**

To kick off 2025, we're offering special rates on all consulting services during January. Whether you're starting new projects or enhancing existing systems, this is the perfect time to bring in extra expertise. Let's make your initiatives a success!

If you have any feedback, questions, or want to discuss a project, just reply to this email—I'd love to hear from you.

Looking forward to continuing this journey together. Here's to your success in 2025!

Best regards,

[Your Name]
[Your Title]

Connect with us:
[Social Media Handle]
[Professional Network]
[Resource Repository]

P.S. If you enjoyed our session, feel free to share your thoughts—we'd love to know what resonates most with you!

*If you no longer wish to receive updates from us, just reply to this email with "Unsubscribe" in the Subject.*
```

## Template Notes

1. Save this template as a Word document (.docx)
2. The `[GivenName]` placeholder will be automatically replaced with each recipient's first name from the CSV
3. Maintain the formatting (bold text, paragraphs) in the Word document
4. You can customize:
   - Company/personal branding
   - Social media links
   - Contact information
   - Special offers and dates
   - Professional titles

## Using the Template

1. Create a new Word document
2. Copy the above content into it
3. Apply desired formatting (fonts, colors, spacing)
4. Save as `Sample.docx`
5. Use with the script:
```powershell
.\sendemailfromoutlook.ps1 -InputTemplate "path\to\Sample.docx" -EmailSubject "Workshop Follow-up" -InputCSV "path\to\recipients.csv"
```
