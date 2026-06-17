#!/usr/bin/env python3
"""
Generate beautiful Islamic Majlis Cards - Part 2: Lesser-Known Teachings
that Muslims commonly underestimate or don't practice correctly.

Creates a visually stunning HTML file with printable cards for sharing at Majlis gatherings.
"""

import os

# Card data: Each card has a topic, icon, colors, Arabic text (if applicable), reference, and explanation
cards_data = [
    {
        "topic": "Istikhara Before Every Decision",
        "icon": "&#x1F4A0;",
        "color_primary": "#1A237E",
        "color_secondary": "#E8EAF6",
        "arabic": "",
        "reference": "Hadith: 'If one of you is concerned about a decision, let him pray two rak'ahs of non-obligatory prayer and then say the Istikhara dua.' (Bukhari)",
        "reminder": "Istikhara is not just for marriage. Use it before any decision - a job, a purchase, travel, even choosing a school. It is asking Allah to choose the best for you."
    },
    {
        "topic": "Adab of the Bathroom",
        "icon": "&#x1F6AA;",
        "color_primary": "#4E342E",
        "color_secondary": "#EFEBE9",
        "arabic": "",
        "reference": "Hadith: 'When entering the toilet, say Bismillah, Allahumma inni a'udhu bika minal khubthi wal khaba'ith. Enter with the left foot and exit with the right foot.' (Bukhari & Muslim)",
        "reminder": "Most people rush in without any dua. The Prophet (SAW) taught specific etiquettes: enter with left foot, exit with right, say the dua, do not face the Qiblah, and do not speak inside."
    },
    {
        "topic": "Silence While Eating",
        "icon": "&#x1F910;",
        "color_primary": "#33691E",
        "color_secondary": "#F1F8E9",
        "arabic": "",
        "reference": "Hadith: The Prophet (SAW) would eat in silence and focus on his food. He (SAW) said: 'Mention the name of Allah and eat with your right hand, and eat from what is nearest to you.' (Bukhari & Muslim)",
        "reminder": "Talking excessively while eating is against the Sunnah. Focus on your food, chew properly, and be mindful of the blessing before you. This also prevents choking and overeating."
    },
    {
        "topic": "Saying Jazakallahu Khairan",
        "icon": "&#x1F4AC;",
        "color_primary": "#0D47A1",
        "color_secondary": "#E3F2FD",
        "arabic": "",
        "reference": "Hadith: 'Whoever has a favor done for him and says Jazak Allahu khairan has done his utmost in praising (thanking) the one who did the favor.' (Tirmidhi)",
        "reminder": "Instead of just saying thanks, say Jazakallahu Khairan (May Allah reward you with good). It is a dua for the person and the best form of gratitude a Muslim can express."
    },
    {
        "topic": "Sin of Delaying Debt Payment",
        "icon": "&#x1F4B3;",
        "color_primary": "#B71C1C",
        "color_secondary": "#FFEBEE",
        "arabic": "",
        "reference": "Hadith: 'Delay in payment by a rich person is injustice (zulm).' (Bukhari & Muslim) | 'The soul of the believer is held back by his debt until it is paid off.' (Tirmidhi)",
        "reminder": "If you owe someone money and have the means to pay, delaying it is a sin. Debts can prevent a person from entering Jannah. Pay your debts before anything else."
    },
    {
        "topic": "Prohibition of Envy (Hasad)",
        "icon": "&#x1F525;",
        "color_primary": "#880E4F",
        "color_secondary": "#FCE4EC",
        "arabic": "",
        "reference": "Hadith: 'Beware of envy, for envy devours good deeds just as fire devours wood.' (Abu Dawud) | 'Do not envy one another, do not hate one another, do not turn away from one another.' (Muslim)",
        "reminder": "Hasad is wishing that a blessing be removed from someone else. It eats your good deeds like fire eats wood. Say MashaAllah, make dua for them, and be content with Allah's decree."
    },
    {
        "topic": "Adab of Yawning",
        "icon": "&#x1F971;",
        "color_primary": "#4A148C",
        "color_secondary": "#F3E5F5",
        "arabic": "",
        "reference": "Hadith: 'Yawning is from Shaytan. If one of you yawns, let him suppress it as much as possible, and let him not say Ha Ha (make a sound), for that is from Shaytan and Shaytan laughs at him.' (Bukhari & Muslim)",
        "reminder": "Cover your mouth when yawning, try to suppress it, and never make a loud sound. It is from Shaytan, and saying A'udhu billah and suppressing it is the Sunnah."
    },
    {
        "topic": "Trimming Nails Within 40 Days",
        "icon": "&#x270B;",
        "color_primary": "#006064",
        "color_secondary": "#E0F7FA",
        "arabic": "",
        "reference": "Hadith: Anas (RA) said: 'The Messenger of Allah (SAW) set a time limit for us for trimming the moustache, clipping the nails, plucking armpit hair, and shaving pubic hair - not to leave it for more than 40 days.' (Muslim)",
        "reminder": "Letting nails grow beyond 40 days is explicitly prohibited. It is part of the Fitrah (natural disposition). The Prophet (SAW) was strict about this hygiene practice."
    },
    {
        "topic": "Prohibition of Sleeping on Stomach",
        "icon": "&#x1F6CF;",
        "color_primary": "#E65100",
        "color_secondary": "#FFF3E0",
        "arabic": "",
        "reference": "Hadith: 'The Prophet (SAW) passed by a man lying on his stomach and nudged him with his foot, saying: This is a way of lying that Allah dislikes.' (Abu Dawud, Tirmidhi)",
        "reminder": "Many people sleep face down out of habit. The Prophet (SAW) explicitly said this sleeping position is disliked by Allah. Sleep on your right side following the Sunnah."
    },
    {
        "topic": "Reading Surah Al-Kahf on Friday",
        "icon": "&#x1F4D6;",
        "color_primary": "#1B5E20",
        "color_secondary": "#E8F5E9",
        "arabic": "",
        "reference": "Hadith: 'Whoever reads Surah Al-Kahf on Friday, a light will shine for him between the two Fridays.' (Al-Hakim, Bayhaqi) | 'He will be given light between the two Fridays.' (Nasai)",
        "reminder": "This simple weekly practice is neglected by most. Read Surah Kahf every Friday (from Maghrib Thursday to Maghrib Friday). It is a shield of light and protection from Dajjal."
    },
    {
        "topic": "Istighfar as a Solution to Problems",
        "icon": "&#x1F333;",
        "color_primary": "#2E7D32",
        "color_secondary": "#C8E6C9",
        "arabic": "",
        "reference": "Quran: 'Ask forgiveness of your Lord. Indeed, He is ever a Perpetual Forgiver. He will send rain to you in abundance, give you increase in wealth and children, and provide for you gardens and rivers.' (Surah Nuh 71:10-12)",
        "reminder": "Istighfar is not just a dhikr. It is a solution to poverty, infertility, drought, anxiety, and hardship. Increase it daily - at minimum 100 times - and watch doors open."
    },
    {
        "topic": "Prohibition of Wasting Water",
        "icon": "&#x1F4A7;",
        "color_primary": "#01579B",
        "color_secondary": "#E1F5FE",
        "arabic": "",
        "reference": "Hadith: 'The Prophet (SAW) passed by Sa'd while he was performing wudu and said: What is this extravagance? Sa'd said: Can there be extravagance in wudu? He (SAW) said: Yes, even if you are at a flowing river.' (Ibn Majah)",
        "reminder": "The Prophet (SAW) made wudu with just one mudd (about 750ml) of water. Even at a river, wasting water is prohibited. Turn off the tap while soaping hands, use only what is needed."
    },
    {
        "topic": "Making Dua for Others in Their Absence",
        "icon": "&#x1F91E;",
        "color_primary": "#311B92",
        "color_secondary": "#EDE7F6",
        "arabic": "",
        "reference": "Hadith: 'The supplication of a Muslim for his brother in his absence is answered. At his head is an angel appointed, and every time he makes dua for his brother with good, the angel says: Ameen, and the same for you.' (Muslim)",
        "reminder": "When you make dua for someone without them knowing, an angel says Ameen and makes the same dua for YOU. This is one of the fastest ways to have your own prayers answered."
    },
    {
        "topic": "Rights of the Road (Removing Harm)",
        "icon": "&#x1F6E4;",
        "color_primary": "#F57F17",
        "color_secondary": "#FFFDE7",
        "arabic": "",
        "reference": "Hadith: 'Removing harmful things from the road is an act of charity (Sadaqah).' (Bukhari & Muslim) | 'Faith has over sixty branches, the highest is La ilaha illallah, and the lowest is removing harm from the path.' (Muslim)",
        "reminder": "If you see a stone, glass, trash, or anything harmful on the road, remove it. This simple act that takes seconds is counted as charity and is a branch of Iman."
    },
    {
        "topic": "Not Pointing at People or Mocking",
        "icon": "&#x1F645;",
        "color_primary": "#C62828",
        "color_secondary": "#FFCDD2",
        "arabic": "",
        "reference": "Quran: 'O you who have believed, let not a people ridicule another people; perhaps they may be better than them.' (Surah Al-Hujurat 49:11) | Hadith: 'It is enough evil for a person to hold his Muslim brother in contempt.' (Muslim)",
        "reminder": "Pointing at someone, laughing at their appearance, mocking their speech or walk - all are forbidden. You do not know whose status is higher with Allah. Ridicule destroys brotherhood."
    },
    {
        "topic": "Breaking Ties of Kinship",
        "icon": "&#x1F517;",
        "color_primary": "#37474F",
        "color_secondary": "#ECEFF1",
        "arabic": "",
        "reference": "Hadith: 'The one who severs ties of kinship will not enter Paradise.' (Bukhari & Muslim) | Quran: 'Would you then, if you were given authority, do mischief in the land, and sever your ties of kinship?' (Surah Muhammad 47:22)",
        "reminder": "Not speaking to relatives, ignoring family members, holding grudges against kin - these can block you from Jannah. Maintain ties even if they cut you off. Visit, call, or message them."
    },
    {
        "topic": "Using the Miswak",
        "icon": "&#x1FAB5;",
        "color_primary": "#795548",
        "color_secondary": "#D7CCC8",
        "arabic": "",
        "reference": "Hadith: 'If it were not that I would be making things too hard for my Ummah, I would have commanded them to use the Siwak (Miswak) before every prayer.' (Bukhari & Muslim)",
        "reminder": "The Prophet (SAW) loved the Miswak so much he almost made it obligatory. Use it before prayers, before entering the house, when waking up, and before reciting Quran. It pleases Allah."
    },
    {
        "topic": "Saying Alhamdulillah After Eating",
        "icon": "&#x1F372;",
        "color_primary": "#FF6F00",
        "color_secondary": "#FFF8E1",
        "arabic": "",
        "reference": "Hadith: 'Allah is pleased with a slave who eats something and praises Him (says Alhamdulillah) for it, or drinks something and praises Him for it.' (Muslim)",
        "reminder": "Most finish eating and immediately get up. Take a moment to say Alhamdulillah al-ladhi at'amani hadha wa razaqanihi min ghayri hawlin minni wa la quwwah. This simple act pleases Allah immensely."
    },
    {
        "topic": "Adab of Sleeping (Dusting the Bed)",
        "icon": "&#x1F303;",
        "color_primary": "#283593",
        "color_secondary": "#C5CAE9",
        "arabic": "",
        "reference": "Hadith: 'When any one of you goes to bed, let him dust off his bed with the inside of his lower garment three times, for he does not know what came onto it after he left it.' (Bukhari & Muslim)",
        "reminder": "Before sleeping: dust your bed 3 times with the edge of your garment (say Bismillah), sleep on your right side, place your right hand under your cheek, and recite the sleeping adhkar."
    },
    {
        "topic": "The Sin of Kibr (Arrogance)",
        "icon": "&#x26A0;",
        "color_primary": "#D32F2F",
        "color_secondary": "#FFCDD2",
        "arabic": "",
        "reference": "Hadith: 'No one who has an atom\\'s weight of arrogance in his heart will enter Paradise.' A man said: What if he likes his clothes and shoes to look good? He (SAW) said: 'Allah is Beautiful and loves beauty. Arrogance means rejecting the truth and looking down on people.' (Muslim)",
        "reminder": "Kibr is not about dressing well. It is rejecting truth when it comes to you and considering yourself above others. Even an atom\\'s weight blocks Jannah. Check yourself daily."
    },
    {
        "topic": "Importance of Niyyah (Intention)",
        "icon": "&#x1F3AF;",
        "color_primary": "#004D40",
        "color_secondary": "#E0F2F1",
        "arabic": "",
        "reference": "Hadith: 'Actions are judged by intentions, and every person will get what they intended.' (Bukhari & Muslim - First Hadith in Sahih Bukhari)",
        "reminder": "Sleeping can be worship if you intend to rest for Fajr. Eating can be worship if you intend to gain strength for ibadah. Renew your intention before EVERY action - work, study, even exercise."
    },
    {
        "topic": "Prohibition of Wasting Food",
        "icon": "&#x1F35E;",
        "color_primary": "#5D4037",
        "color_secondary": "#EFEBE9",
        "arabic": "",
        "reference": "Quran: 'Eat and drink, but do not waste. Indeed, He does not like those who waste.' (Surah Al-A'raf 7:31) | Hadith: 'If a morsel of food falls, pick it up, clean it, and eat it.' (Muslim)",
        "reminder": "Take only what you can finish. Do not leave food on your plate. If food falls, pick it up and eat it. The Prophet (SAW) would lick his fingers and clean the plate. Wasting food is hated by Allah."
    },
    {
        "topic": "Sitting While Drinking Water",
        "icon": "&#x1F4A7;",
        "color_primary": "#0277BD",
        "color_secondary": "#B3E5FC",
        "arabic": "",
        "reference": "Hadith: 'The Prophet (SAW) forbade drinking while standing.' (Muslim) | In another narration: 'Had the one who drank standing known what was in his stomach, he would have vomited.' (Ahmad)",
        "reminder": "Drinking while standing is against the Sunnah and medically harmful. Sit down, say Bismillah, drink in 3 sips, breathe outside the cup between sips, and say Alhamdulillah when done."
    },
    {
        "topic": "Attending Janazah (Funeral Prayer)",
        "icon": "&#x1F54B;",
        "color_primary": "#263238",
        "color_secondary": "#CFD8DC",
        "arabic": "",
        "reference": "Hadith: 'Whoever attends the funeral prayer will have one Qirat of reward. Whoever stays until the burial will have two Qirats. Each Qirat is like Mount Uhud.' (Bukhari & Muslim)",
        "reminder": "Two Qirats - each the size of Mount Uhud in reward - just for attending a funeral and burial. Yet most Muslims skip it unless it is someone very close. It is a right of the deceased upon the community."
    },
    {
        "topic": "Being Mindful of What Enters Your Ears",
        "icon": "&#x1F442;",
        "color_primary": "#1A237E",
        "color_secondary": "#E8EAF6",
        "arabic": "",
        "reference": "Quran: 'Indeed, the hearing, the sight, and the heart - about all those one will be questioned.' (Surah Al-Isra 17:36) | 'And of the people is he who buys the amusement of speech to mislead from the way of Allah.' (Surah Luqman 31:6)",
        "reminder": "Your ears will testify on the Day of Judgment about what you listened to. Be intentional about what enters your ears - podcasts, conversations, background noise. Fill them with Quran and beneficial knowledge."
    },
    {
        "topic": "Being Just Even with Enemies",
        "icon": "&#x2696;",
        "color_primary": "#4A148C",
        "color_secondary": "#F3E5F5",
        "arabic": "",
        "reference": "Quran: 'O you who have believed, be persistently standing firm for Allah as witnesses in justice, and do not let the hatred of a people prevent you from being just. Be just; that is nearer to righteousness.' (Surah Al-Ma'idah 5:8)",
        "reminder": "Islam demands justice even toward those who hate you. Do not lie about your enemy, do not exaggerate their faults, do not deny their rights. Justice is non-negotiable regardless of personal feelings."
    },
    {
        "topic": "Speak Good or Remain Silent",
        "icon": "&#x1F910;",
        "color_primary": "#006064",
        "color_secondary": "#E0F7FA",
        "arabic": "",
        "reference": "Hadith: 'Whoever believes in Allah and the Last Day, let him speak good or remain silent.' (Bukhari & Muslim)",
        "reminder": "Before speaking, ask: Is it true? Is it beneficial? Is it the right time? If it fails any test, stay silent. Most sins come from the tongue. Silence is safety, and good speech is rewarded."
    },
]


def generate_html():
    """Generate beautiful HTML cards."""
    
    html_content = """<!DOCTYPE html>
<html lang="en" dir="ltr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Islamic Majlis Cards Part 2 - Lesser-Known Teachings We Underestimate</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Amiri:wght@400;700&family=Poppins:wght@300;400;500;600;700&display=swap');
        
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Poppins', sans-serif;
            background: linear-gradient(135deg, #1a0033 0%, #4a0072 50%, #7b1fa2 100%);
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
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.25);
        }
        
        .header h1 {
            font-size: 2.5em;
            color: #4A148C;
            margin-bottom: 10px;
            font-weight: 700;
        }
        
        .header .arabic-title {
            font-family: 'Amiri', serif;
            font-size: 2.2em;
            color: #6A1B9A;
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
        
        .header .subtitle {
            display: inline-block;
            background: linear-gradient(135deg, #4A148C, #7B1FA2);
            color: white;
            padding: 8px 20px;
            border-radius: 20px;
            font-size: 0.9em;
            margin-top: 15px;
            font-weight: 500;
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
        <div class="arabic-title">\u0628\u0650\u0633\u0652\u0645\u0650 \u0627\u0644\u0644\u0651\u064e\u0647\u0650 \u0627\u0644\u0631\u0651\u064e\u062d\u0652\u0645\u064e\u0646\u0650 \u0627\u0644\u0631\u0651\u064e\u062d\u0650\u064a\u0645\u0650</div>
        <h1>\u2728 Islamic Majlis Cards - Part 2 \u2728</h1>
        <p><strong>Lesser-Known Teachings We Underestimate</strong><br>
        These are the practices and prohibitions that many Muslims overlook or don't think about in daily life. They are authentic from the Quran and Sunnah, yet commonly neglected.</p>
        <span class="subtitle">27 Overlooked Teachings Every Muslim Should Know</span>
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
        <p><strong>May Allah help us practice these forgotten Sunnahs and avoid what He dislikes.</strong></p>
        <p>Share with others - many Muslims are unaware of these teachings.</p>
        <div class="dua">\u0631\u064e\u0628\u0651\u064e\u0646\u064e\u0627 \u062a\u064e\u0642\u064e\u0628\u0651\u064e\u0644\u0652 \u0645\u0650\u0646\u0651\u064e\u0627 \u0625\u0650\u0646\u0651\u064e\u0643\u064e \u0623\u064e\u0646\u062a\u064e \u0627\u0644\u0633\u0651\u064e\u0645\u0650\u064a\u0639\u064f \u0627\u0644\u0652\u0639\u064e\u0644\u0650\u064a\u0645\u064f</div>
        <p style="margin-top: 15px; opacity: 0.7; font-size: 0.85em;">Prepared for Islamic Majlis | Part 2: Things We Overlook</p>
    </div>
</body>
</html>
"""
    return html_content


def main():
    """Main function to generate the Majlis cards Part 2."""
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Islamic_Majlis_Cards_Part2.html")
    
    html = generate_html()
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"Islamic Majlis Cards Part 2 generated successfully!")
    print(f"Output: {output_path}")
    print(f"Total cards: {len(cards_data)}")
    print(f"\nTopics covered:")
    for i, card in enumerate(cards_data, 1):
        print(f"  {i}. {card['topic']}")
    print(f"\nOpen the HTML file in any browser to view the cards.")
    print(f"You can also print them directly from the browser (Ctrl+P).")


if __name__ == "__main__":
    main()
