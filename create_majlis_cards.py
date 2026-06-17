#!/usr/bin/env python3
"""
Generate beautiful Islamic Majlis Cards - Basic teachings from Quran and Hadith
that Muslims often underestimate or overlook in daily life.

Creates a visually stunning HTML file with printable cards for sharing at Majlis gatherings.
"""

import os

# Card data: Each card has a topic, Arabic text (if applicable), reference, and explanation
cards_data = [
    {
        "topic": "Saying Bismillah",
        "icon": "&#x1F4BF;",
        "color_primary": "#1B5E20",
        "color_secondary": "#E8F5E9",
        "arabic": "\u0628\u0650\u0633\u0652\u0645\u0650 \u0627\u0644\u0644\u0651\u064E\u0647\u0650 \u0627\u0644\u0631\u0651\u064E\u062D\u0652\u0645\u064E\u0646\u0650 \u0627\u0644\u0631\u0651\u064E\u062D\u0650\u064A\u0645\u0650",
        "reference": "Hadith: The Prophet (SAW) said: 'Any matter of importance not begun with Bismillah is devoid of blessing.' (Abu Dawud)",
        "reminder": "Begin every action - eating, drinking, entering home, starting work - with Bismillah. It invites Allah's blessings into everything you do."
    },
    {
        "topic": "Spreading Salam",
        "icon": "&#x1F91D;",
        "color_primary": "#0D47A1",
        "color_secondary": "#E3F2FD",
        "arabic": "\u0627\u0644\u0633\u0651\u064E\u0644\u0627\u0645\u064F \u0639\u064E\u0644\u064E\u064A\u0652\u0643\u064F\u0645\u0652 \u0648\u064E\u0631\u064E\u062D\u0652\u0645\u064E\u0629\u064F \u0627\u0644\u0644\u0651\u064E\u0647\u0650 \u0648\u064E\u0628\u064E\u0631\u064E\u0643\u064E\u0627\u062A\u064F\u0647\u064F",
        "reference": "Hadith: 'You will not enter Paradise until you believe, and you will not believe until you love one another. Shall I tell you something that if you do it, you will love one another? Spread Salam among yourselves.' (Muslim)",
        "reminder": "Say Salam to those you know and those you don't. It is a prayer of peace and a means of building love in the Ummah."
    },
    {
        "topic": "Rights of Neighbors",
        "icon": "&#x1F3E0;",
        "color_primary": "#4A148C",
        "color_secondary": "#F3E5F5",
        "arabic": "",
        "reference": "Quran: 'Worship Allah and associate nothing with Him, and to parents do good, and to relatives, orphans, the needy, the near neighbor, the neighbor farther away...' (Surah An-Nisa 4:36)",
        "reminder": "The Prophet (SAW) said Jibreel kept advising him about neighbors until he thought they would be made heirs. Check on them, share food, avoid harming them."
    },
    {
        "topic": "Honesty & Truthfulness",
        "icon": "&#x2696;",
        "color_primary": "#B71C1C",
        "color_secondary": "#FFEBEE",
        "arabic": "",
        "reference": "Hadith: 'Truthfulness leads to righteousness, and righteousness leads to Paradise. A man keeps speaking the truth until he is recorded with Allah as a truthful person (Siddiq).' (Bukhari & Muslim)",
        "reminder": "Avoid lies even in jest. The Prophet (SAW) never lied, even in difficult situations. Truthfulness is the foundation of a Muslim's character."
    },
    {
        "topic": "Kindness to Parents",
        "icon": "&#x1F495;",
        "color_primary": "#880E4F",
        "color_secondary": "#FCE4EC",
        "arabic": "",
        "reference": "Quran: 'And your Lord has decreed that you not worship except Him, and to parents, good treatment. Whether one or both of them reach old age, do not say to them \"uff\" and do not repel them.' (Surah Al-Isra 17:23)",
        "reminder": "Even a sigh of frustration (uff) towards parents is prohibited. Serve them with love, patience, and gratitude. Their pleasure is Allah's pleasure."
    },
    {
        "topic": "Avoiding Backbiting",
        "icon": "&#x1F6AB;",
        "color_primary": "#E65100",
        "color_secondary": "#FFF3E0",
        "arabic": "",
        "reference": "Quran: 'Do not spy or backbite each other. Would one of you like to eat the flesh of his dead brother? You would detest it.' (Surah Al-Hujurat 49:12)",
        "reminder": "Backbiting (Gheebah) is mentioning something about your brother that he would dislike, even if it is true. Guard your tongue as you guard your prayers."
    },
    {
        "topic": "Being Grateful (Shukr)",
        "icon": "&#x1F64F;",
        "color_primary": "#F57F17",
        "color_secondary": "#FFFDE7",
        "arabic": "",
        "reference": "Quran: 'If you are grateful, I will surely increase you [in favor]; but if you deny, indeed, My punishment is severe.' (Surah Ibrahim 14:7)",
        "reminder": "Say Alhamdulillah often. Be grateful for health, family, food, and faith. Gratitude increases blessings; ingratitude invites loss."
    },
    {
        "topic": "Remembrance of Allah (Dhikr)",
        "icon": "&#x1F4FF;",
        "color_primary": "#006064",
        "color_secondary": "#E0F7FA",
        "arabic": "",
        "reference": "Quran: 'Verily, in the remembrance of Allah do hearts find rest.' (Surah Ar-Ra'd 13:28)",
        "reminder": "Keep your tongue moist with dhikr - SubhanAllah, Alhamdulillah, Allahu Akbar, Astaghfirullah. It is light on the tongue but heavy on the scales."
    },
    {
        "topic": "Punctuality in Salah",
        "icon": "&#x1F54C;",
        "color_primary": "#1A237E",
        "color_secondary": "#E8EAF6",
        "arabic": "",
        "reference": "Quran: 'Indeed, prayer has been decreed upon the believers a decree of specified times.' (Surah An-Nisa 4:103)",
        "reminder": "The first thing we will be asked about on the Day of Judgment is Salah. Pray on time, with concentration (Khushu). It is the pillar of your Deen."
    },
    {
        "topic": "Controlling Anger",
        "icon": "&#x1F9D8;",
        "color_primary": "#311B92",
        "color_secondary": "#EDE7F6",
        "arabic": "",
        "reference": "Hadith: 'The strong man is not the one who can wrestle, but the strong man is the one who controls himself at the time of anger.' (Bukhari & Muslim)",
        "reminder": "When angry: say A'udhu billahi min ash-shaytan ir-rajeem, sit down if standing, make wudu, and remain silent. Anger is from Shaytan."
    },
    {
        "topic": "Charity (Sadaqah)",
        "icon": "&#x1F4B0;",
        "color_primary": "#1B5E20",
        "color_secondary": "#E8F5E9",
        "arabic": "",
        "reference": "Hadith: 'Charity does not decrease wealth. No one forgives another except that Allah increases his honor.' (Muslim)",
        "reminder": "Give charity daily, even if small. A smile is charity. Removing harm from the road is charity. Even a kind word is Sadaqah."
    },
    {
        "topic": "Good Manners (Akhlaq)",
        "icon": "&#x1F31F;",
        "color_primary": "#FF6F00",
        "color_secondary": "#FFF8E1",
        "arabic": "",
        "reference": "Hadith: 'The most beloved of you to me and the closest to me in position on the Day of Resurrection are those of you who are best in character.' (Tirmidhi)",
        "reminder": "Good character outweighs extra prayers and fasting. Be gentle, speak kindly, show mercy, and treat all people with respect."
    },
    {
        "topic": "Eating with Right Hand",
        "icon": "&#x1F37D;",
        "color_primary": "#33691E",
        "color_secondary": "#F1F8E9",
        "arabic": "",
        "reference": "Hadith: 'When one of you eats, let him eat with his right hand, and when he drinks, let him drink with his right hand.' (Muslim)",
        "reminder": "A simple Sunnah we often neglect. Eat with the right hand, say Bismillah, eat from what is nearest to you, and do not waste food."
    },
    {
        "topic": "Replying to a Sneeze",
        "icon": "&#x1F927;",
        "color_primary": "#01579B",
        "color_secondary": "#E1F5FE",
        "arabic": "",
        "reference": "Hadith: 'When one of you sneezes and says Alhamdulillah, say to him Yarhamuk-Allah (May Allah have mercy on you).' (Bukhari)",
        "reminder": "Saying Alhamdulillah after sneezing and responding Yarhamuk-Allah is a forgotten Sunnah that builds bonds of mercy in our community."
    },
    {
        "topic": "Avoiding Suspicion & Spying",
        "icon": "&#x1F440;",
        "color_primary": "#4E342E",
        "color_secondary": "#EFEBE9",
        "arabic": "",
        "reference": "Quran: 'O you who have believed, avoid much [negative] assumption. Indeed, some assumption is sin. And do not spy or backbite each other.' (Surah Al-Hujurat 49:12)",
        "reminder": "Do not assume the worst about your Muslim brother or sister. Give 70 excuses before judging. Protect their honor as you would protect your own."
    },
    {
        "topic": "Making Dua (Supplication)",
        "icon": "&#x1F932;",
        "color_primary": "#004D40",
        "color_secondary": "#E0F2F1",
        "arabic": "",
        "reference": "Hadith: 'Dua is the essence of worship.' (Tirmidhi) | Quran: 'Call upon Me; I will respond to you.' (Surah Ghafir 40:60)",
        "reminder": "Never underestimate dua. Ask Allah for everything - big or small. The best times: last third of night, between Adhan and Iqamah, while prostrating."
    },
    {
        "topic": "Visiting the Sick",
        "icon": "&#x1F3E5;",
        "color_primary": "#C62828",
        "color_secondary": "#FFCDD2",
        "arabic": "",
        "reference": "Hadith: 'When a Muslim visits a sick Muslim in the morning, 70,000 angels pray for him until evening, and if he visits in the evening, 70,000 angels pray for him until morning.' (Tirmidhi)",
        "reminder": "Visiting the sick is a right upon every Muslim. It brings immense reward and reminds us of Allah's blessings of health."
    },
    {
        "topic": "Keeping Promises",
        "icon": "&#x1F91D;",
        "color_primary": "#283593",
        "color_secondary": "#C5CAE9",
        "arabic": "",
        "reference": "Quran: 'O you who have believed, fulfill [all] contracts.' (Surah Al-Ma'idah 5:1) | Hadith: 'The signs of a hypocrite are three: when he speaks he lies, when he promises he breaks it, when he is trusted he betrays.' (Bukhari & Muslim)",
        "reminder": "A Muslim's word is his bond. Do not make promises you cannot keep. Breaking promises is a sign of hypocrisy (Nifaq)."
    },
    {
        "topic": "Lowering the Gaze",
        "icon": "&#x1F441;",
        "color_primary": "#37474F",
        "color_secondary": "#ECEFF1",
        "arabic": "",
        "reference": "Quran: 'Tell the believing men to lower their gaze and guard their modesty. That is purer for them.' (Surah An-Nur 24:30)",
        "reminder": "The first glance is forgiven, but the second is a sin. Guard your eyes in public, on screens, and in gatherings. It protects the heart."
    },
    {
        "topic": "Sleeping & Waking Sunnah",
        "icon": "&#x1F319;",
        "color_primary": "#1A237E",
        "color_secondary": "#E8EAF6",
        "arabic": "",
        "reference": "Hadith: 'When you go to bed, perform wudu as for prayer, then lie on your right side and say: O Allah, I surrender myself to You...' (Bukhari & Muslim)",
        "reminder": "Sleep on your right side, recite Ayatul Kursi, the last 2 verses of Al-Baqarah, and the 3 Quls. Wake for Tahajjud if you can - it is the best prayer after Fard."
    },
]


def generate_html():
    """Generate beautiful HTML cards."""
    
    html_content = """<!DOCTYPE html>
<html lang="en" dir="ltr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Islamic Majlis Cards - Basic Teachings We Often Underestimate</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Amiri:wght@400;700&family=Poppins:wght@300;400;500;600;700&display=swap');
        
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Poppins', sans-serif;
            background: linear-gradient(135deg, #0f3443 0%, #34e89e 100%);
            min-height: 100vh;
            padding: 40px 20px;
        }
        
        .header {
            text-align: center;
            margin-bottom: 50px;
            padding: 40px 20px;
            background: rgba(255, 255, 255, 0.95);
            border-radius: 20px;
            max-width: 900px;
            margin-left: auto;
            margin-right: auto;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.15);
        }
        
        .header h1 {
            font-size: 2.5em;
            color: #1B5E20;
            margin-bottom: 10px;
            font-weight: 700;
        }
        
        .header .arabic-title {
            font-family: 'Amiri', serif;
            font-size: 2.2em;
            color: #2E7D32;
            margin-bottom: 15px;
            direction: rtl;
        }
        
        .header p {
            font-size: 1.1em;
            color: #555;
            max-width: 700px;
            margin: 0 auto;
            line-height: 1.6;
        }
        
        .cards-container {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(380px, 1fr));
            gap: 30px;
            max-width: 1400px;
            margin: 0 auto;
            padding: 20px;
        }
        
        .card {
            background: white;
            border-radius: 16px;
            overflow: hidden;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.12);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
            position: relative;
        }
        
        .card:hover {
            transform: translateY(-5px);
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.2);
        }
        
        .card-header {
            padding: 25px 25px 15px;
            position: relative;
        }
        
        .card-icon {
            font-size: 2.5em;
            margin-bottom: 10px;
            display: inline-block;
        }
        
        .card-number {
            position: absolute;
            top: 15px;
            right: 20px;
            width: 35px;
            height: 35px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 700;
            font-size: 0.85em;
            color: white;
        }
        
        .card-title {
            font-size: 1.4em;
            font-weight: 700;
            margin-bottom: 5px;
        }
        
        .card-arabic {
            font-family: 'Amiri', serif;
            font-size: 1.5em;
            direction: rtl;
            text-align: right;
            padding: 15px 25px;
            margin: 0 20px;
            border-radius: 10px;
            line-height: 1.8;
        }
        
        .card-body {
            padding: 20px 25px 25px;
        }
        
        .card-reference {
            font-size: 0.88em;
            line-height: 1.6;
            padding: 12px 15px;
            border-radius: 8px;
            margin-bottom: 15px;
            border-left: 4px solid;
            font-style: italic;
        }
        
        .card-reminder {
            font-size: 0.95em;
            line-height: 1.7;
            color: #333;
            padding-top: 10px;
            border-top: 1px dashed #ddd;
        }
        
        .card-reminder::before {
            content: "\\1F4A1 ";
        }
        
        .footer {
            text-align: center;
            margin-top: 50px;
            padding: 30px;
            color: white;
            font-size: 0.95em;
        }
        
        .footer p {
            margin-bottom: 8px;
            opacity: 0.9;
        }
        
        .footer .dua {
            font-family: 'Amiri', serif;
            font-size: 1.4em;
            margin-top: 15px;
            direction: rtl;
        }
        
        @media print {
            body {
                background: white;
                padding: 10px;
            }
            .card {
                break-inside: avoid;
                box-shadow: none;
                border: 2px solid #ddd;
                margin-bottom: 20px;
            }
            .card:hover {
                transform: none;
            }
            .cards-container {
                grid-template-columns: repeat(2, 1fr);
                gap: 15px;
            }
        }
        
        @media (max-width: 480px) {
            .cards-container {
                grid-template-columns: 1fr;
            }
            .header h1 {
                font-size: 1.8em;
            }
        }
    </style>
</head>
<body>
    <div class="header">
        <div class="arabic-title">\u0628\u0650\u0633\u0652\u0645\u0650 \u0627\u0644\u0644\u0651\u064E\u0647\u0650 \u0627\u0644\u0631\u0651\u064E\u062D\u0652\u0645\u064E\u0646\u0650 \u0627\u0644\u0631\u0651\u064E\u062D\u0650\u064A\u0645\u0650</div>
        <h1>\u2728 Islamic Majlis Cards \u2728</h1>
        <p><strong>Basic Teachings We Often Underestimate</strong><br>
        A reminder of the simple yet powerful practices from the Quran and Sunnah that we overlook in our daily lives. Let us revive these teachings and earn the pleasure of Allah (SWT).</p>
    </div>
    
    <div class="cards-container">
"""

    # Generate each card
    for i, card in enumerate(cards_data, 1):
        arabic_section = ""
        if card["arabic"]:
            arabic_section = f"""
        <div class="card-arabic" style="background: {card['color_secondary']}; color: {card['color_primary']};">
            {card['arabic']}
        </div>"""
        
        html_content += f"""
        <div class="card">
            <div class="card-header" style="background: {card['color_secondary']};">
                <span class="card-icon">{card['icon']}</span>
                <span class="card-number" style="background: {card['color_primary']};">{i}</span>
                <h2 class="card-title" style="color: {card['color_primary']};">{card['topic']}</h2>
            </div>{arabic_section}
            <div class="card-body">
                <div class="card-reference" style="background: {card['color_secondary']}40; border-color: {card['color_primary']};">
                    {card['reference']}
                </div>
                <div class="card-reminder">
                    {card['reminder']}
                </div>
            </div>
        </div>
"""

    html_content += """
    </div>
    
    <div class="footer">
        <p><strong>May Allah guide us all to act upon these teachings.</strong></p>
        <p>Share with others - spreading knowledge is Sadaqah Jariyah (continuous charity).</p>
        <div class="dua">\u0627\u0644\u0644\u0651\u064E\u0647\u064F\u0645\u0651\u064E \u0627\u062C\u0652\u0639\u064E\u0644\u0652\u0646\u064E\u0627 \u0645\u0650\u0646\u064E \u0627\u0644\u0651\u064E\u0630\u0650\u064A\u0646\u064E \u064A\u064E\u0633\u0652\u062A\u064E\u0645\u0650\u0639\u064F\u0648\u0646\u064E \u0627\u0644\u0652\u0642\u064E\u0648\u0652\u0644\u064E \u0641\u064E\u064A\u064E\u062A\u0651\u064E\u0628\u0650\u0639\u064F\u0648\u0646\u064E \u0623\u064E\u062D\u0652\u0633\u064E\u0646\u064E\u0647\u064F</div>
        <p style="margin-top: 15px; opacity: 0.7; font-size: 0.85em;">Prepared for Islamic Majlis | Designed with love for the Ummah</p>
    </div>
</body>
</html>
"""
    return html_content


def main():
    """Main function to generate the Majlis cards."""
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Islamic_Majlis_Cards.html")
    
    html = generate_html()
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"Islamic Majlis Cards generated successfully!")
    print(f"Output: {output_path}")
    print(f"Total cards: {len(cards_data)}")
    print(f"\nTopics covered:")
    for i, card in enumerate(cards_data, 1):
        print(f"  {i}. {card['topic']}")
    print(f"\nOpen the HTML file in any browser to view the cards.")
    print(f"You can also print them directly from the browser (Ctrl+P).")


if __name__ == "__main__":
    main()
