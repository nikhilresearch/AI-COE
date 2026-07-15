import os
import subprocess

slides = {
    1: "Welcome to the AI Build Bootcamp and Hackathon 2026 at BIMTECH. This flagship event, held on 7th and 8th August 2026, is organized in collaboration with Reskilll. We are here to empower future managers with real, hands-on AI skills. Artificial Intelligence is transforming every business function — from Finance and Marketing to HR, Operations, and Strategy. Tomorrow's managers will not only use AI, but build AI-enabled business solutions. Our objective is to equip every first-year MBA student with hands-on experience in designing and building AI-powered business applications.",
    2: "Why does this matter to our students? AI is rapidly becoming a core management competency. Through this bootcamp and hackathon, students will learn to convert business problems into AI solutions, work collaboratively in multidisciplinary teams, build AI-enabled prototypes without extensive coding, develop innovation and design thinking skills, and present solutions before industry experts. The outcomes are truly exciting — hands-on AI exposure, practical problem-solving, real industry interaction, an innovation mindset, and portfolio-worthy projects. As we say: Learn AI by Building, Not Just by Reading.",
    3: "Let me walk you through the event structure. On Day 1, the 7th of August, students will participate in the AI Bootcamp. This includes an introduction to AI tools, guided hands-on sessions, team formation, and problem statement briefing. On Day 2, the 8th of August, the AI Hackathon begins — students develop prototypes with mentoring support, make project submissions, and get evaluated by Reskilll. Then comes the exciting Pitch Day on 19th August, where the Top 10 teams present before industry experts, advisory committee, and faculty jury. The Top 3 teams will be specially rewarded!",
    4: "Programme Chairs play a critical role in making this initiative truly successful. We need your support in encouraging maximum student participation, ensuring students attend as per batch allocation, minimizing class disruptions during the event, motivating students to treat the hackathon as a genuine learning opportunity, and coordinating with faculty mentors wherever required. In return, students will receive a joint BIMTECH and Reskilll participation certificate, industry exposure, an AI portfolio project, recognition for top teams, and an incredible opportunity to present before industry leaders.",
    5: "This is more than just a hackathon. It is the first flagship initiative of the AI Centre of Excellence at BIMTECH. Our vision is to create AI-literate management professionals capable of solving real business challenges using AI. Our journey is clear: from the AI Bootcamp, to the AI Hackathon, to the Pitch Day, to the formal launch of our AI Centre of Excellence, then to industry projects, and ultimately building a full AI Innovation Ecosystem at BIMTECH. Together, we are shaping the future of management education in India. Thank you."
}

edge_tts_path = r"C:\Users\admin\AppData\Local\Python\pythoncore-3.14-64\Scripts\edge-tts.exe"
voice = "en-IN-PrabhatNeural"

for i, text in slides.items():
    output_file = f"slide_{i}.mp3"
    print(f"Generating {output_file}...")
    subprocess.run([edge_tts_path, "--voice", voice, "--text", text, "--write-media", output_file], check=True)

print("All audio files generated successfully!")
