const pptInput = document.querySelector("#pptInput");
const docInput = document.querySelector("#docInput");
const voiceSelect = document.querySelector("#voiceSelect");
const prevBtn = document.querySelector("#prevBtn");
const playBtn = document.querySelector("#playBtn");
const nextBtn = document.querySelector("#nextBtn");
const slideCounter = document.querySelector("#slideCounter");
const statusText = document.querySelector("#statusText");
const slideCanvas = document.querySelector("#slideCanvas");
const scriptText = document.querySelector("#scriptText");
const speech = window.speechSynthesis;

let slides = [];
let docSegments = [];
let docFullText = "";
let currentSlide = 0;
let isReading = false;
let advanceTimer;
let voices = [];

function setStatus(message) {
  statusText.textContent = message;
}

function textFromXml(xmlText) {
  const xml = new DOMParser().parseFromString(xmlText, "application/xml");
  return [...xml.getElementsByTagName("*")]
    .filter((node) => node.localName === "t")
    .map((node) => node.textContent.trim())
    .filter(Boolean)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

function splitTextIntoSegments(text, count) {
  const paragraphs = text
    .split(/\n{2,}/)
    .map((part) => part.replace(/\s+/g, " ").trim())
    .filter(Boolean);

  if (!count || paragraphs.length === count) {
    return paragraphs;
  }

  const sentences = text
    .replace(/\s+/g, " ")
    .split(/(?<=[.!?؟։…])\s+/)
    .map((part) => part.trim())
    .filter(Boolean);

  if (sentences.length <= count) {
    return sentences;
  }

  const bucketSize = Math.ceil(sentences.length / count);
  return Array.from({ length: count }, (_, index) =>
    sentences.slice(index * bucketSize, (index + 1) * bucketSize).join(" ")
  ).filter(Boolean);
}

async function readZipFile(file) {
  if (!window.JSZip) {
    throw new Error("JSZip ვერ ჩაიტვირთა. შეამოწმე ინტერნეტი და თავიდან სცადე.");
  }

  const buffer = await file.arrayBuffer();
  return window.JSZip.loadAsync(buffer);
}

async function loadPowerPoint(file) {
  const zip = await readZipFile(file);
  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => Number(a.match(/slide(\d+)\.xml/)[1]) - Number(b.match(/slide(\d+)\.xml/)[1]));

  const parsedSlides = [];

  for (const name of slideFiles) {
    const xml = await zip.file(name).async("text");
    const text = textFromXml(xml);
    parsedSlides.push({
      title: text.split(/[.!?۔։]/)[0]?.trim() || `სლაიდი ${parsedSlides.length + 1}`,
      text: text || "ამ სლაიდზე ტექსტი ვერ მოიძებნა.",
    });
  }

  slides = parsedSlides;
  if (docFullText) {
    docSegments = splitTextIntoSegments(docFullText, slides.length);
  }
  currentSlide = 0;
  setStatus(`${slides.length} სლაიდი ჩაიტვირთა`);
  renderSlide();
}

async function loadWord(file) {
  const zip = await readZipFile(file);
  const documentFile = zip.file("word/document.xml");
  if (!documentFile) {
    throw new Error("Word document.xml ვერ მოიძებნა");
  }

  const documentXml = await documentFile.async("text");
  const xml = new DOMParser().parseFromString(documentXml, "application/xml");
  const paragraphs = [...xml.getElementsByTagName("*")]
    .filter((node) => node.localName === "p")
    .map((paragraph) =>
      [...paragraph.getElementsByTagName("*")]
        .filter((node) => node.localName === "t")
        .map((node) => node.textContent)
        .join("")
        .replace(/\s+/g, " ")
        .trim()
    )
    .filter(Boolean);

  docFullText = paragraphs.join("\n\n");
  docSegments = splitTextIntoSegments(docFullText, slides.length);
  setStatus(`Word ტექსტი ჩაიტვირთა: ${docSegments.length} ნაწილი`);
  renderSlide();
}

function getSlideScript(index) {
  return docSegments[index] || slides[index]?.text || "";
}

function renderSlide() {
  const slide = slides[currentSlide];
  const count = slides.length;

  slideCounter.textContent = `სლაიდი ${count ? currentSlide + 1 : 0} / ${count}`;
  prevBtn.disabled = !count || currentSlide === 0;
  nextBtn.disabled = !count || currentSlide === count - 1;
  playBtn.disabled = !count || !speech;

  if (!slide) {
    slideCanvas.innerHTML = `
      <h1>ქართული პრეზენტაციის მკითხველი</h1>
      <p>ატვირთე .pptx პრეზენტაცია და სურვილის შემთხვევაში .docx ტექსტი. შემდეგ დააჭირე „წაკითხვა“-ს.</p>
    `;
    scriptText.textContent = "აქ გამოჩნდება ტექსტი, რომელსაც ქართული ხმა წაიკითხავს.";
    return;
  }

  const script = getSlideScript(currentSlide);
  slideCanvas.innerHTML = `
    <h1>${escapeHtml(slide.title)}</h1>
    <p>${escapeHtml(slide.text)}</p>
  `;
  scriptText.textContent = script || "ამ სლაიდისთვის წასაკითხი ტექსტი არ მოიძებნა.";
}

function escapeHtml(text) {
  return text.replace(/[&<>"']/g, (char) => {
    const entities = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#039;",
    };
    return entities[char];
  });
}

function populateVoices() {
  if (!speech) {
    voiceSelect.innerHTML = '<option value="">ხმა მიუწვდომელია</option>';
    playBtn.disabled = true;
    return;
  }

  voices = speech.getVoices();
  const selectedValue = voiceSelect.value;
  const georgianVoices = voices.filter((voice) => voice.lang?.toLowerCase().startsWith("ka"));
  const preferredVoices = georgianVoices.length ? georgianVoices : voices;

  voiceSelect.innerHTML = "";
  for (const voice of preferredVoices) {
    const option = document.createElement("option");
    option.value = voice.name;
    option.textContent = `${voice.name} (${voice.lang})`;
    voiceSelect.append(option);
  }

  if (!preferredVoices.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "ხმა ვერ მოიძებნა";
    voiceSelect.append(option);
  }

  if (selectedValue) {
    voiceSelect.value = selectedValue;
  }
}

function selectedVoice() {
  return voices.find((voice) => voice.name === voiceSelect.value) || voices.find((voice) => voice.lang?.toLowerCase().startsWith("ka"));
}

function stopReading() {
  clearTimeout(advanceTimer);
  speech?.cancel();
  isReading = false;
  document.body.classList.remove("reading");
  playBtn.textContent = "წაკითხვა";
}

function readCurrentSlide() {
  if (!speech) {
    setStatus("ამ ბრაუზერს ხმით წაკითხვა არ შეუძლია");
    return;
  }

  const script = getSlideScript(currentSlide);

  if (!script) {
    setStatus("ამ სლაიდის ტექსტი ცარიელია");
    return;
  }

  clearTimeout(advanceTimer);
  speech.cancel();

  const utterance = new SpeechSynthesisUtterance(script);
  utterance.lang = "ka-GE";
  utterance.rate = 0.92;
  utterance.pitch = 1;

  const voice = selectedVoice();
  if (voice) {
    utterance.voice = voice;
  }

  utterance.onstart = () => {
    isReading = true;
    document.body.classList.add("reading");
    playBtn.textContent = "შეჩერება";
    setStatus("ქართული ხმა კითხულობს");
  };

  utterance.onend = () => {
    advanceTimer = setTimeout(() => {
      if (currentSlide < slides.length - 1) {
        currentSlide += 1;
        renderSlide();
        readCurrentSlide();
        return;
      }

      stopReading();
      setStatus("პრეზენტაცია დასრულდა");
    }, 2000);
  };

  utterance.onerror = () => {
    stopReading();
    setStatus("ხმის წაკითხვა ვერ მოხერხდა. სცადე Chrome ან Edge.");
  };

  speech.speak(utterance);
}

function goToSlide(nextIndex) {
  if (!slides.length) {
    return;
  }

  currentSlide = Math.max(0, Math.min(slides.length - 1, nextIndex));
  renderSlide();

  if (isReading) {
    readCurrentSlide();
  }
}

pptInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) {
    return;
  }

  stopReading();
  setStatus("PowerPoint იტვირთება...");

  try {
    await loadPowerPoint(file);
  } catch (error) {
    setStatus(error.message || "PowerPoint ვერ ჩაიტვირთა");
  }
});

docInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) {
    return;
  }

  stopReading();
  setStatus("Word ტექსტი იტვირთება...");

  try {
    await loadWord(file);
  } catch (error) {
    setStatus("Word ტექსტი ვერ ჩაიტვირთა");
  }
});

playBtn.addEventListener("click", () => {
  if (isReading) {
    stopReading();
    setStatus("წაკითხვა შეჩერდა");
    return;
  }

  readCurrentSlide();
});

prevBtn.addEventListener("click", () => goToSlide(currentSlide - 1));
nextBtn.addEventListener("click", () => goToSlide(currentSlide + 1));

speech?.addEventListener("voiceschanged", populateVoices);
populateVoices();
renderSlide();
