/* ============================================================
   Conversion Heaven – Application Logic
   ============================================================ */

(function () {
    "use strict";

    /* ---------- DOM References ---------- */
    const themeToggle = document.getElementById("themeToggle");
    const themeIcon = document.getElementById("themeIcon");
    const loadingOverlay = document.getElementById("loadingOverlay");
    const thankYouModal = document.getElementById("thankYouModal");
    const modalCloseBtn = document.getElementById("modalCloseBtn");
    const particleCanvas = document.getElementById("particleCanvas");
    const ctx = particleCanvas.getContext("2d");

    /* ---------- Theme Toggle ---------- */
    function applyTheme(theme) {
        document.documentElement.setAttribute("data-theme", theme);
        themeIcon.textContent = theme === "dark" ? "🌙" : "☀️";
        localStorage.setItem("ch-theme", theme);
    }

    themeToggle.addEventListener("click", function () {
        var current = document.documentElement.getAttribute("data-theme");
        applyTheme(current === "dark" ? "light" : "dark");
    });

    var stored = localStorage.getItem("ch-theme");
    if (stored) applyTheme(stored);

    /* ---------- Particle System ---------- */
    var particles = [];
    var EMOJIS = ["🎉", "✨", "🥳", "🎊", "⭐", "💜", "🔥", "🚀"];

    function resizeCanvas() {
        particleCanvas.width = window.innerWidth;
        particleCanvas.height = window.innerHeight;
    }
    window.addEventListener("resize", resizeCanvas);
    resizeCanvas();

    function spawnParticles(x, y) {
        for (var i = 0; i < 30; i++) {
            particles.push({
                x: x,
                y: y,
                vx: (Math.random() - 0.5) * 12,
                vy: (Math.random() - 0.5) * 12 - 4,
                life: 1,
                emoji: EMOJIS[Math.floor(Math.random() * EMOJIS.length)],
                size: 16 + Math.random() * 18,
                rotation: Math.random() * 360,
                rotationSpeed: (Math.random() - 0.5) * 10
            });
        }
    }

    function animateParticles() {
        ctx.clearRect(0, 0, particleCanvas.width, particleCanvas.height);
        for (var i = particles.length - 1; i >= 0; i--) {
            var p = particles[i];
            p.x += p.vx;
            p.y += p.vy;
            p.vy += 0.25;
            p.life -= 0.018;
            p.rotation += p.rotationSpeed;
            if (p.life <= 0) {
                particles.splice(i, 1);
                continue;
            }
            ctx.save();
            ctx.globalAlpha = p.life;
            ctx.translate(p.x, p.y);
            ctx.rotate((p.rotation * Math.PI) / 180);
            ctx.font = p.size + "px serif";
            ctx.textAlign = "center";
            ctx.textBaseline = "middle";
            ctx.fillText(p.emoji, 0, 0);
            ctx.restore();
        }
        requestAnimationFrame(animateParticles);
    }
    animateParticles();

    /* ---------- Loading helpers ---------- */
    function showLoading() {
        loadingOverlay.classList.add("active");
    }
    function hideLoading() {
        loadingOverlay.classList.remove("active");
    }

    /* ---------- Thank-you modal ---------- */
    function showThankYou() {
        thankYouModal.classList.add("active");
    }
    modalCloseBtn.addEventListener("click", function () {
        thankYouModal.classList.remove("active");
    });

    /* ---------- Download helper ---------- */
    function triggerDownload(blob, filename, btnEvent) {
        var url = URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        /* Revoke immediately – "delete from database after conversion" */
        setTimeout(function () { URL.revokeObjectURL(url); }, 500);

        /* Emoji particles at click position */
        if (btnEvent) {
            spawnParticles(btnEvent.clientX, btnEvent.clientY);
        } else {
            spawnParticles(window.innerWidth / 2, window.innerHeight / 2);
        }

        /* Show thank-you modal after a brief pause */
        setTimeout(showThankYou, 600);
    }

    /* ---------- Helper: read file as ArrayBuffer ---------- */
    function readFileAsArrayBuffer(file) {
        return new Promise(function (resolve, reject) {
            var reader = new FileReader();
            reader.onload = function () { resolve(reader.result); };
            reader.onerror = function () { reject(reader.error); };
            reader.readAsArrayBuffer(file);
        });
    }

    /* ---------- Helper: read file as Text ---------- */
    function readFileAsText(file) {
        return new Promise(function (resolve, reject) {
            var reader = new FileReader();
            reader.onload = function () { resolve(reader.result); };
            reader.onerror = function () { reject(reader.error); };
            reader.readAsText(file);
        });
    }

    /* ---------- Helper: read file as Data URL ---------- */
    function readFileAsDataURL(file) {
        return new Promise(function (resolve, reject) {
            var reader = new FileReader();
            reader.onload = function () { resolve(reader.result); };
            reader.onerror = function () { reject(reader.error); };
            reader.readAsDataURL(file);
        });
    }

    /* ---------- Library availability check ---------- */
    function checkLibraries(names) {
        var missing = [];
        var libs = {
            "jspdf": function () { return window.jspdf && window.jspdf.jsPDF; },
            "mammoth": function () { return typeof mammoth !== "undefined"; },
            "htmlDocx": function () { return window.htmlDocx && window.htmlDocx.asBlob; },
            "XLSX": function () { return typeof XLSX !== "undefined"; },
            "pdfjsLib": function () { return typeof pdfjsLib !== "undefined"; },
            "JSZip": function () { return typeof JSZip !== "undefined"; },
            "html2canvas": function () { return typeof html2canvas !== "undefined"; },
            "lamejs": function () { return typeof lamejs !== "undefined"; }
        };
        for (var i = 0; i < names.length; i++) {
            if (libs[names[i]] && !libs[names[i]]()) {
                missing.push(names[i]);
            }
        }
        if (missing.length > 0) {
            throw new Error(
                "Required libraries failed to load: " + missing.join(", ") +
                ". Please check your internet connection and reload the page."
            );
        }
    }

    /* ---------- Conversion: Word → PDF ---------- */
    async function convertWordToPdf(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["mammoth", "jspdf", "html2canvas"]);
            var arrayBuffer = await readFileAsArrayBuffer(file);
            var result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
            var html = result.value;

            /* Render HTML in a hidden container so html2canvas can capture it */
            var container = document.createElement("div");
            container.style.cssText = "position:fixed;left:-9999px;top:0;width:595px;padding:40px;background:#fff;color:#000;font-family:serif;font-size:12pt;line-height:1.6;";
            container.innerHTML = html;
            document.body.appendChild(container);

            var canvas = await html2canvas(container, { scale: 2, useCORS: true });
            document.body.removeChild(container);

            var imgData = canvas.toDataURL("image/jpeg", 0.95);
            var jsPDF = window.jspdf.jsPDF;
            var pageW = 595.28; /* A4 width in pt */
            var pageH = 841.89; /* A4 height in pt */
            var imgW = pageW;
            var imgH = (canvas.height * pageW) / canvas.width;

            var doc = new jsPDF({ unit: "pt", format: "a4" });
            var position = 0;

            /* Paginate the captured image across multiple PDF pages */
            while (position < imgH) {
                if (position > 0) doc.addPage();
                doc.addImage(imgData, "JPEG", 0, -position, imgW, imgH);
                position += pageH;
            }

            var blob = doc.output("blob");
            triggerDownload(blob, file.name.replace(/\.docx?$/i, "") + ".pdf", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: PPT → PDF ---------- */
    async function convertPptToPdf(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["jspdf", "JSZip", "html2canvas"]);
            var arrayBuffer = await readFileAsArrayBuffer(file);
            var jsPDF = window.jspdf.jsPDF;

            /* PPTX is a ZIP containing XML slides */
            var zip = await JSZip.loadAsync(arrayBuffer);
            var slideTexts = [];

            /* Gather slide file names and sort them */
            var slideFiles = [];
            zip.forEach(function (relativePath) {
                if (/^ppt\/slides\/slide\d+\.xml$/i.test(relativePath)) {
                    slideFiles.push(relativePath);
                }
            });
            slideFiles.sort(function (a, b) {
                var matchA = a.match(/slide(\d+)/i);
                var matchB = b.match(/slide(\d+)/i);
                var numA = matchA ? parseInt(matchA[1], 10) : 0;
                var numB = matchB ? parseInt(matchB[1], 10) : 0;
                return numA - numB;
            });

            for (var s = 0; s < slideFiles.length; s++) {
                var xml = await zip.file(slideFiles[s]).async("string");
                /* Parse XML and extract text from <a:t> elements (DrawingML text) */
                var parser = new DOMParser();
                var xmlDoc = parser.parseFromString(xml, "application/xml");
                var textNodes = xmlDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "t");
                var parts = [];
                for (var t = 0; t < textNodes.length; t++) {
                    var content = textNodes[t].textContent || "";
                    if (content.trim()) parts.push(content.trim());
                }
                slideTexts.push(parts);
            }

            if (slideTexts.length === 0) {
                slideTexts.push(["(No slide content could be extracted from " + file.name + ")"]);
            }

            /* Build an HTML representation of each slide and render via html2canvas */
            var doc = new jsPDF({ unit: "pt", format: [720, 540], orientation: "landscape" });
            var pageW = 720;
            var pageH = 540;

            for (var i = 0; i < slideTexts.length; i++) {
                if (i > 0) doc.addPage([720, 540], "landscape");

                var container = document.createElement("div");
                container.style.cssText = "position:fixed;left:-9999px;top:0;width:720px;height:540px;background:#fff;color:#000;font-family:sans-serif;display:flex;flex-direction:column;justify-content:center;padding:40px;box-sizing:border-box;";

                var titleEl = document.createElement("div");
                titleEl.style.cssText = "font-size:24px;font-weight:bold;margin-bottom:20px;color:#333;";
                titleEl.textContent = "Slide " + (i + 1);
                container.appendChild(titleEl);

                var textParts = slideTexts[i];
                for (var p = 0; p < textParts.length; p++) {
                    var line = document.createElement("div");
                    line.style.cssText = "font-size:14px;margin-bottom:8px;color:#222;line-height:1.5;";
                    line.textContent = textParts[p];
                    container.appendChild(line);
                }

                document.body.appendChild(container);
                var canvas = await html2canvas(container, { scale: 2, useCORS: true, width: 720, height: 540 });
                document.body.removeChild(container);

                var imgData = canvas.toDataURL("image/jpeg", 0.95);
                doc.addImage(imgData, "JPEG", 0, 0, pageW, pageH);
            }

            var blob = doc.output("blob");
            triggerDownload(blob, file.name.replace(/\.pptx?$/i, "") + ".pdf", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: Excel → PDF ---------- */
    async function convertExcelToPdf(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["jspdf", "XLSX"]);
            var arrayBuffer = await readFileAsArrayBuffer(file);
            var workbook = XLSX.read(arrayBuffer, { type: "array" });

            var jsPDF = window.jspdf.jsPDF;
            var doc = new jsPDF({ unit: "pt", format: "a4" });

            var y = 40;
            var pageHeight = doc.internal.pageSize.getHeight();

            workbook.SheetNames.forEach(function (name, idx) {
                if (idx > 0) {
                    doc.addPage();
                    y = 40;
                }
                doc.setFontSize(14);
                doc.text("Sheet: " + name, 40, y);
                y += 24;
                doc.setFontSize(10);

                var sheet = workbook.Sheets[name];
                var text = XLSX.utils.sheet_to_csv(sheet);
                var lines = doc.splitTextToSize(text, 500);
                for (var i = 0; i < lines.length; i++) {
                    if (y > pageHeight - 40) {
                        doc.addPage();
                        y = 40;
                    }
                    doc.text(lines[i], 40, y);
                    y += 14;
                }
            });

            var blob = doc.output("blob");
            triggerDownload(blob, file.name.replace(/\.xlsx?$/i, "") + ".pdf", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: Image → PDF ---------- */
    async function convertImageToPdf(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["jspdf"]);

            /* Only allow PNG and JPEG */
            var fileType = (file.type || "").toLowerCase();
            if (fileType !== "image/png" && fileType !== "image/jpeg") {
                throw new Error("Only PNG and JPEG images are supported for Image to PDF conversion.");
            }

            var dataUrl = await readFileAsDataURL(file);
            var jsPDF = window.jspdf.jsPDF;
            var doc = new jsPDF({ unit: "pt", format: "a4" });

            var img = new Image();
            await new Promise(function (resolve, reject) {
                img.onload = resolve;
                img.onerror = reject;
                img.src = dataUrl;
            });

            var pageW = doc.internal.pageSize.getWidth() - 80;
            var pageH = doc.internal.pageSize.getHeight() - 80;
            var ratio = Math.min(pageW / img.width, pageH / img.height, 1);
            var w = img.width * ratio;
            var h = img.height * ratio;
            var format = fileType === "image/png" ? "PNG" : "JPEG";
            doc.addImage(dataUrl, format, 40, 40, w, h);

            var blob = doc.output("blob");
            triggerDownload(blob, file.name.replace(/\.[^.]+$/, "") + ".pdf", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: PDF → Image ---------- */
    async function convertPdfToImage(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["pdfjsLib", "JSZip"]);
            var arrayBuffer = await readFileAsArrayBuffer(file);
            var pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            var numPages = pdf.numPages;

            if (numPages === 1) {
                /* Single page - download as single PNG */
                var page = await pdf.getPage(1);
                var scale = 2;
                var viewport = page.getViewport({ scale: scale });
                var canvas = document.createElement("canvas");
                canvas.width = viewport.width;
                canvas.height = viewport.height;
                var context = canvas.getContext("2d");
                await page.render({ canvasContext: context, viewport: viewport }).promise;

                await new Promise(function (resolve) {
                    canvas.toBlob(function (blob) {
                        triggerDownload(blob, file.name.replace(/\.pdf$/i, "") + ".png", btnEvent);
                        resolve();
                    }, "image/png");
                });
            } else {
                /* Multiple pages - create ZIP file with all images */
                var zip = new JSZip();
                var baseName = file.name.replace(/\.pdf$/i, "");

                for (var i = 1; i <= numPages; i++) {
                    var page = await pdf.getPage(i);
                    var scale = 2;
                    var viewport = page.getViewport({ scale: scale });
                    var canvas = document.createElement("canvas");
                    canvas.width = viewport.width;
                    canvas.height = viewport.height;
                    var context = canvas.getContext("2d");
                    await page.render({ canvasContext: context, viewport: viewport }).promise;

                    /* Convert canvas to blob */
                    var blob = await new Promise(function (resolve) {
                        canvas.toBlob(resolve, "image/png");
                    });

                    /* Add to ZIP with page number */
                    zip.file(baseName + "_page_" + i + ".png", blob);
                }

                /* Generate ZIP file */
                var zipBlob = await zip.generateAsync({ type: "blob" });
                triggerDownload(zipBlob, baseName + "_images.zip", btnEvent);
            }
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: TXT → Word ---------- */
    async function convertTxtToWord(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["htmlDocx"]);
            var text = await readFileAsText(file);
            
            /* Convert plain text to HTML with proper formatting */
            var htmlContent = text
                .split('\n')
                .map(function(line) {
                    /* Escape HTML special characters */
                    var escaped = line
                        .replace(/&/g, '&amp;')
                        .replace(/</g, '&lt;')
                        .replace(/>/g, '&gt;')
                        .replace(/"/g, '&quot;')
                        .replace(/'/g, '&#039;');
                    return '<p>' + (escaped || '&nbsp;') + '</p>';
                })
                .join('');
            
            var htmlDoc = '<html><head><meta charset="utf-8"></head><body>' + htmlContent + '</body></html>';
            var converted = window.htmlDocx.asBlob(htmlDoc);
            triggerDownload(converted, file.name.replace(/\.txt$/i, "") + ".docx", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: Video → Audio (offline, always MP3) ---------- */
    async function convertVideoToAudio(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["lamejs"]);
            var arrayBuffer = await readFileAsArrayBuffer(file);
            var audioCtx = new (window.AudioContext || window.webkitAudioContext)();
            var audioBuffer = await audioCtx.decodeAudioData(arrayBuffer);
            audioCtx.close();

            /* Convert decoded audio to WAV, then encode to MP3 with lamejs */
            var numChannels = audioBuffer.numberOfChannels;
            var sampleRate = audioBuffer.sampleRate;
            var channels = [];
            for (var c = 0; c < numChannels; c++) {
                channels.push(audioBuffer.getChannelData(c));
            }

            /* Convert float samples to 16-bit PCM */
            var numSamples = audioBuffer.length;
            var left = new Int16Array(numSamples);
            var right = numChannels > 1 ? new Int16Array(numSamples) : left;
            for (var i = 0; i < numSamples; i++) {
                var s = Math.max(-1, Math.min(1, channels[0][i]));
                left[i] = s < 0 ? s * 0x8000 : s * 0x7FFF;
                if (numChannels > 1) {
                    s = Math.max(-1, Math.min(1, channels[1][i]));
                    right[i] = s < 0 ? s * 0x8000 : s * 0x7FFF;
                }
            }

            /* Encode to MP3 using lamejs */
            var mp3enc = new lamejs.Mp3Encoder(numChannels > 1 ? 2 : 1, sampleRate, 128);
            var mp3Data = [];
            var blockSize = 1152;
            for (var i = 0; i < numSamples; i += blockSize) {
                var leftChunk = left.subarray(i, i + blockSize);
                var rightChunk = numChannels > 1 ? right.subarray(i, i + blockSize) : leftChunk;
                var mp3buf = numChannels > 1
                    ? mp3enc.encodeBuffer(leftChunk, rightChunk)
                    : mp3enc.encodeBuffer(leftChunk);
                if (mp3buf.length > 0) mp3Data.push(mp3buf);
            }
            var end = mp3enc.flush();
            if (end.length > 0) mp3Data.push(end);

            var blob = new Blob(mp3Data, { type: "audio/mpeg" });
            triggerDownload(blob, file.name.replace(/\.[^.]+$/, "") + ".mp3", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion dispatcher ---------- */
    var CONVERTERS = {
        "word-to-pdf": convertWordToPdf,
        "ppt-to-pdf": convertPptToPdf,
        "excel-to-pdf": convertExcelToPdf,
        "image-to-pdf": convertImageToPdf,
        "pdf-to-image": convertPdfToImage,
        "txt-to-word": convertTxtToWord,
        "video-to-audio": convertVideoToAudio
    };

    /* ---------- Wire up cards ---------- */
    document.querySelectorAll(".card").forEach(function (card) {
        var type = card.getAttribute("data-conversion");
        var input = card.querySelector("input[type=file]");
        var fileNameSpan = card.querySelector(".file-name");
        var convertBtn = card.querySelector(".convert-btn");
        var dropZone = card.querySelector(".drop-zone");
        var selectedFile = null;

        function selectFile(file) {
            selectedFile = file;
            fileNameSpan.textContent = file.name;
            convertBtn.disabled = false;
        }

        input.addEventListener("change", function () {
            if (input.files && input.files.length > 0) {
                selectFile(input.files[0]);
            }
        });

        /* ---------- Drag & Drop ---------- */
        if (dropZone) {
            var dragCounter = 0;
            var acceptAttr = input.getAttribute("accept") || "";

            function isFileAccepted(file) {
                if (!acceptAttr) return true;
                var accepts = acceptAttr.split(",").map(function (s) { return s.trim().toLowerCase(); });
                var fileName = file.name.toLowerCase();
                var fileType = (file.type || "").toLowerCase();
                return accepts.some(function (a) {
                    if (a.startsWith(".")) {
                        return fileName.endsWith(a);
                    }
                    if (a.endsWith("/*")) {
                        return fileType.startsWith(a.replace("/*", "/"));
                    }
                    return fileType === a;
                });
            }

            dropZone.addEventListener("dragover", function (e) {
                e.preventDefault();
                e.stopPropagation();
            });

            dropZone.addEventListener("dragenter", function (e) {
                e.preventDefault();
                e.stopPropagation();
                dragCounter++;
                dropZone.classList.add("drag-over");
            });

            dropZone.addEventListener("dragleave", function (e) {
                e.preventDefault();
                e.stopPropagation();
                dragCounter--;
                if (dragCounter <= 0) {
                    dragCounter = 0;
                    dropZone.classList.remove("drag-over");
                }
            });

            dropZone.addEventListener("drop", function (e) {
                e.preventDefault();
                e.stopPropagation();
                dragCounter = 0;
                dropZone.classList.remove("drag-over");
                if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                    var droppedFile = e.dataTransfer.files[0];
                    if (isFileAccepted(droppedFile)) {
                        selectFile(droppedFile);
                    } else {
                        alert("Invalid file type. Please drop a supported file (" + acceptAttr + ").");
                    }
                }
            });
        }

        convertBtn.addEventListener("click", function (e) {
            if (!selectedFile) return;
            var converter = CONVERTERS[type];
            if (converter) {
                converter(selectedFile, e);
                /* Clear file reference after conversion starts ("delete from database") */
                setTimeout(function () {
                    selectedFile = null;
                    input.value = "";
                    fileNameSpan.textContent = "";
                    convertBtn.disabled = true;
                }, 200);
            }
        });
    });

})();
