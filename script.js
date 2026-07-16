const pptInput = document.querySelector("#pptInput");
const docInput = document.querySelector("#docInput");
const voiceSelect = document.querySelector("#voiceSelect");
const pptName = document.querySelector("#pptName");
const docName = document.querySelector("#docName");
const prevBtn = document.querySelector("#prevBtn");
const playBtn = document.querySelector("#playBtn");
const nextBtn = document.querySelector("#nextBtn");
const testVoiceBtn = document.querySelector("#testVoiceBtn");
const diagnosticText = document.querySelector("#diagnosticText");
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
let readingParts = [];
let readingPartIndex = 0;
let isVisualFallback = false;
let activeAudio;

function setStatus(message) {
  statusText.textContent = message;
}

function setDiagnostic(message) {
  diagnosticText.textContent = message;
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

function createSlidesFromDocSegments() {
  slides = docSegments.map((segment, index) => ({
    title: `ნაწილი ${index + 1}`,
    text: segment,
  }));
  currentSlide = 0;
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
  if (!slides.length) {
    createSlidesFromDocSegments();
  }
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
  playBtn.disabled = !count;

  if (!slide) {
    slideCanvas.innerHTML = `
      <h2>დაიწყე ფაილების ატვირთვით</h2>
      <p>ატვირთე .pptx პრეზენტაცია და სურვილის შემთხვევაში .docx ტექსტი. შემდეგ დააჭირე „წაკითხვა“-ს.</p>
    `;
    scriptText.textContent = "აქ გამოჩნდება ტექსტი, რომელსაც ქართული ხმა წაიკითხავს.";
    return;
  }

  const script = getSlideScript(currentSlide);
  slideCanvas.innerHTML = `
    <h2>${escapeHtml(slide.title)}</h2>
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
    testVoiceBtn.disabled = true;
    setDiagnostic("ამ ბრაუზერში speechSynthesis საერთოდ არ არის. გახსენი Chrome ან Edge-ში.");
    return;
  }

  voices = speech.getVoices();
  const selectedValue = voiceSelect.value;
  const georgianVoices = voices.filter((voice) => voice.lang?.toLowerCase().startsWith("ka"));
  const preferredVoices = [...georgianVoices, ...voices.filter((voice) => !voice.lang?.toLowerCase().startsWith("ka"))];

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

  if (!voices.length) {
    setDiagnostic("ხმების სია ჯერ ცარიელია. დააჭირე „ხმის ტესტი“-ს ან გახსენი გვერდი Chrome/Edge-ში.");
    return;
  }

  if (georgianVoices.length) {
    setDiagnostic(`ნაპოვნია ${voices.length} ხმა. ქართული ხმა მზადაა: ${georgianVoices[0].name}.`);
  } else {
    setDiagnostic(`ნაპოვნია ${voices.length} ხმა, მაგრამ ქართული voice არ ჩანს. სხვა voice შეიძლება ქართულს არ კითხულობდეს.`);
  }
}

function selectedVoice() {
  return voices.find((voice) => voice.name === voiceSelect.value) || voices.find((voice) => voice.lang?.toLowerCase().startsWith("ka")) || null;
}

function stopReading() {
  clearTimeout(advanceTimer);
  speech?.cancel();
  if (activeAudio) {
    activeAudio.pause();
    activeAudio.src = "";
  }
  activeAudio = null;
  readingParts = [];
  readingPartIndex = 0;
  isReading = false;
  isVisualFallback = false;
  document.body.classList.remove("reading");
  playBtn.textContent = "წაკითხვა";
}

async function speakWithBackend(text, onDone) {
  const result = await fetch("/api/tts", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ text }),
  });

  if (!result.ok) {
    const message = await result.text();
    throw new Error(message || "TTS backend failed");
  }

  const audioBlob = await result.blob();
  const audioUrl = URL.createObjectURL(audioBlob);
  const audio = new Audio(audioUrl);
  activeAudio = audio;

  audio.onplay = () => onDone?.("start");
  audio.onended = () => {
    URL.revokeObjectURL(audioUrl);
    if (activeAudio === audio) {
      activeAudio = null;
    }
    onDone?.("end");
  };
  audio.onerror = () => {
    URL.revokeObjectURL(audioUrl);
    if (activeAudio === audio) {
      activeAudio = null;
    }
    onDone?.("fail", "audio playback failed");
  };

  await audio.play();
}

function speakText(text, onDone) {
  speakWithBackend(text, onDone).catch(() => {
    speakWithBrowser(text, onDone);
  });
}

function speakWithBrowser(text, onDone) {
  if (!speech) {
    onDone?.("fail", "ამ ბრაუზერს ხმით წაკითხვა არ შეუძლია");
    return;
  }

  const utterance = new SpeechSynthesisUtterance(text);
  const voice = selectedVoice();
  utterance.lang = voice?.lang || "ka-GE";
  utterance.rate = 0.88;
  utterance.pitch = 1;

  if (voice) {
    utterance.voice = voice;
  }

  let started = false;
  let startedAt = 0;
  const startGuard = setTimeout(() => {
    if (!started) {
      setStatus("ხმა არ დაიწყო. სცადე Chrome/Edge ან სხვა voice dropdown-იდან.");
      onDone?.("fail", "ხმა არ დაიწყო");
    }
  }, 1400);

  utterance.onstart = () => {
    started = true;
    startedAt = performance.now();
    clearTimeout(startGuard);
    onDone?.("start");
  };

  utterance.onend = () => {
    clearTimeout(startGuard);
    if (startedAt && performance.now() - startedAt < 250) {
      onDone?.("fail", "ხმა მაშინვე დასრულდა");
      return;
    }
    onDone?.("end");
  };

  utterance.onerror = (event) => {
    clearTimeout(startGuard);
    stopReading();
    setStatus(`ხმის შეცდომა: ${event.error || "unknown"}. სცადე სხვა ბრაუზერი ან voice.`);
    onDone?.("fail", event.error || "unknown");
  };

  speech.speak(utterance);
}

function estimatedReadingTime(text) {
  const words = text.split(/\s+/).filter(Boolean).length;
  return Math.max(2200, Math.min(12000, words * 420));
}

function visualFallbackReadCurrentSlide(reason) {
  const script = getSlideScript(currentSlide);

  isReading = true;
  isVisualFallback = true;
  document.body.classList.add("reading");
  playBtn.textContent = "შეჩერება";
  setStatus(`ხმა არ ირთვება (${reason}). demo რეჟიმში სლაიდი ავტომატურად გადავა.`);

  clearTimeout(advanceTimer);
  advanceTimer = setTimeout(() => {
    finishSlideReading();
  }, estimatedReadingTime(script));
}

function splitForSpeech(text) {
  const sentences = text
    .replace(/\s+/g, " ")
    .split(/(?<=[.!?։…])\s+/)
    .map((part) => part.trim())
    .filter(Boolean);

  const parts = [];
  let current = "";

  for (const sentence of sentences.length ? sentences : [text]) {
    if (`${current} ${sentence}`.trim().length > 180 && current) {
      parts.push(current);
      current = sentence;
    } else {
      current = `${current} ${sentence}`.trim();
    }
  }

  if (current) {
    parts.push(current);
  }

  return parts;
}

function finishSlideReading() {
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
}

function speakNextPart() {
  const text = readingParts[readingPartIndex];

  if (isVisualFallback) {
    return;
  }

  if (!text) {
    finishSlideReading();
    return;
  }

  speakText(text, (state) => {
    if (state === "start") {
      isReading = true;
      document.body.classList.add("reading");
      playBtn.textContent = "შეჩერება";
      setStatus(`კითხულობს: ნაწილი ${readingPartIndex + 1} / ${readingParts.length}`);
    }

    if (state === "end") {
      readingPartIndex += 1;
      speakNextPart();
    }

    if (state === "fail") {
      visualFallbackReadCurrentSlide("TTS unavailable");
    }
  });
}

function readCurrentSlide() {
  if (!speech) {
    visualFallbackReadCurrentSlide("speechSynthesis არ არის");
    return;
  }

  const script = getSlideScript(currentSlide);

  if (!script) {
    setStatus("ამ სლაიდის ტექსტი ცარიელია");
    return;
  }

  clearTimeout(advanceTimer);
  speech.cancel();
  readingParts = splitForSpeech(script);
  readingPartIndex = 0;
  speakNextPart();
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

  pptName.textContent = file.name;
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

  docName.textContent = file.name;
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

testVoiceBtn.addEventListener("click", () => {
  speech?.cancel();
  setStatus("ხმის ტესტი იწყება...");
  speakText("გამარჯობა, ეს არის ქართული ხმის ტესტი.", (state, reason) => {
    if (state === "start") {
      setStatus("ხმის ტესტი მუშაობს");
    }

    if (state === "end") {
      setStatus("ხმის ტესტი დასრულდა");
    }

    if (state === "fail") {
      setStatus(`ხმის ტესტი ვერ გავიდა: ${reason}. გახსენი რეალურ Chrome/Edge-ში.`);
    }
  });
});

speech?.addEventListener("voiceschanged", populateVoices);
populateVoices();
setTimeout(populateVoices, 600);
renderSlide();
