document.addEventListener('DOMContentLoaded', () => {
    // Real Data from Excel files
    const tableData = [
        // Penahanan (2).xls
        { id: 1, noDok: '2026-H2.0-7101.0-K.1.1-000007', tgl: '2026-01-22', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Daging Olahan Babi', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Impor' },
        { id: 2, noDok: '2026-H2.0-7101.0-K.1.1-000003', tgl: '2026-01-07', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Daging Ayam', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Impor' },
        { id: 3, noDok: '2026-H2.0-7101.0-K.1.1-000002', tgl: '2026-01-05', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Daging Olahan Babi', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Impor' },
        { id: 4, noDok: '2026-H2.0-7101.0-K.1.1-000001', tgl: '2026-01-01', pengirim: 'NN', penerima: 'NN', asal: 'PHILIPINA', tujuan: 'INDONESIA - KOTA BITUNG', komoditas: 'Unggas Besar Kesayangan', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Impor' },
        { id: 5, noDok: '2026-H2.0-7101.0-K.1.1-000005', tgl: '2026-01-14', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Daging Olahan Sapi', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Impor' },
        { id: 6, noDok: '2026-H2.0-7101.0-K.1.1-000006', tgl: '2026-01-15', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Daging Olahan Babi', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Impor' },
        { id: 7, noDok: '2026-H2.0-7101.0-K.1.1-000004', tgl: '2026-01-13', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Daging Babi', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Impor' },

        // Penahanan (3).xls
        { id: 8, noDok: '2026-I2.0-7101.0-K.1.1-000004', tgl: '2026-01-30', pengirim: 'No Name', penerima: 'No Name', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Ikan Teri', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Ikan', permohonan: 'Impor' },
        { id: 9, noDok: '2026-I2.0-7101.0-K.1.1-000002', tgl: '2026-01-20', pengirim: 'No Name', penerima: 'No Name', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Ikan Patin', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Ikan', permohonan: 'Impor' },
        { id: 10, noDok: '2026-I2.0-7101.0-K.1.1-000001', tgl: '2026-01-09', pengirim: 'No Name', penerima: 'No Name', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Sidat', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Ikan', permohonan: 'Impor' },
        { id: 11, noDok: '2026-I2.0-7101.0-K.1.1-000003', tgl: '2026-01-28', pengirim: 'No Name', penerima: 'No Name', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'Ikan Teri', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Ikan', permohonan: 'Impor' },

        // Penahanan (4).xls
        { id: 12, noDok: '2026-T2.0-7101.0-K.1.1-000004', tgl: '2026-01-20', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'CABE KERING', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Tumbuhan', permohonan: 'Impor' },
        { id: 13, noDok: '2026-T2.0-7101.0-K.1.1-000002', tgl: '2026-01-15', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'JAMUR KERING', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Tumbuhan', permohonan: 'Impor' },
        { id: 14, noDok: '2026-T2.0-7101.0-K.1.1-000001', tgl: '2026-01-06', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'BUAH-BUAHAN', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Tumbuhan', permohonan: 'Impor' },
        { id: 15, noDok: '2026-T2.0-7101.0-K.1.1-000007', tgl: '2026-01-31', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'BAWANG DAUN', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Tumbuhan', permohonan: 'Impor' },
        { id: 16, noDok: '2026-T2.0-7101.0-K.1.1-000006', tgl: '2026-01-22', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'BUAH-BUAHAN', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Tumbuhan', permohonan: 'Impor' },
        { id: 17, noDok: '2026-T2.0-7101.0-K.1.1-000005', tgl: '2026-01-21', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'BIDARA/JUJUBE', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Tumbuhan', permohonan: 'Impor' },
        { id: 18, noDok: '2026-T2.0-7101.1-K.1.1-000001', tgl: '2026-01-05', pengirim: 'NN', penerima: 'NN', asal: 'CHINA', tujuan: 'INDONESIA - KOTA MANADO', komoditas: 'BIJI CHIA', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Tumbuhan', permohonan: 'Impor' },

        // Penahanan (5).xls
        { id: 19, noDok: '2026-H4.0-7103.0-K.1.1-000039', tgl: '2026-01-11', pengirim: 'PINDI KURNIAWAN', penerima: 'AGUNG WAHYU PRIBADI', asal: 'KOTA MANADO', tujuan: 'KEPULAUAN SANGIHE', komoditas: 'Burung Parkit', tkp: 'BKHIT Sulawesi Utara - Pelabuhan Laut Tahuna', ket: 'Lengkap', category: 'Hewan', permohonan: 'Domestik Masuk' },
        { id: 20, noDok: '2026-H4.0-7101.1-K.1.1-000008', tgl: '2026-01-08', pengirim: 'MERLIN MATAHARI', penerima: 'MERLIN MATAHARI', asal: 'KOTA TERNATE', tujuan: 'KOTA MANADO', komoditas: 'Kucing', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Domestik Masuk' },
        { id: 21, noDok: '2026-H4.0-7101.1-K.1.1-000005', tgl: '2026-01-06', pengirim: 'Aldaier Makatindu', penerima: 'Jouce Kansil', asal: 'KEPULAUAN SANGIHE', tujuan: 'KOTA MANADO', komoditas: 'Ayam', tkp: 'BKHIT Sulawesi Utara - Bandara Sam Ratulangi', ket: 'Lengkap', category: 'Hewan', permohonan: 'Domestik Masuk' },
    ];

    const tableBody = document.getElementById('table-body');
    const modal = document.getElementById('detail-modal');
    const closeModalBtn = document.querySelector('.close-modal');
    const previewContainer = document.getElementById('preview-container');
    const fileViewer = document.getElementById('file-viewer');
    const pdfFrame = document.getElementById('pdf-frame');
    const viewerFilename = document.getElementById('viewer-filename');
    const WORKER_URL = 'https://gakkum-monitoring-api.karantinaikanbitung.workers.dev';
    const boxPemusnahan = document.getElementById('box-pemusnahan');
    const boxPenolakan = document.getElementById('box-penolakan');
    let currentUploadContext = null;
    let currentDocId = null; // To track which document we are working on
    let uploadedFiles = {};

    // Render Table
    // Render Table
    function renderTable(data) {
        tableBody.innerHTML = '';
        data.forEach(row => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${row.noDok}</td>
                <td>${row.tgl}</td>
                <td>${row.category}</td>
                <td>${row.permohonan}</td>
                <td>${row.pengirim}</td>
                <td>${row.penerima}</td>
                <td>${row.asal}</td>
                <td>${row.tujuan}</td>
                <td>${row.komoditas}</td>
                <td>${row.tkp}</td>
                <td>${row.ket}</td>
                <td>
                    <button class="btn btn-primary btn-sm" onclick="openDetail(${row.id})">
                        <i class="fa-solid fa-eye"></i> Detail
                    </button>
                </td>
            `;
            tableBody.appendChild(tr);
        });
    }

    // ... (rest of code) ...



    // Modal Logic
    window.openDetail = (id) => {
        currentDocId = id;
        localStorage.setItem('activeDocId', id);
        modal.style.display = 'flex';
        resetView();
        checkExistingFiles(id);
    };

    async function checkExistingFiles(id) {
        const contexts = ['LI', 'Kronologi', 'PUL BAKET', 'Penolakan', 'Pemusnahan'];
        const docIdAtStart = id;

        for (const ctx of contexts) {
            try {
                // If user closed modal or switched doc, stop checking
                if (currentDocId !== docIdAtStart) return;

                const response = await fetch(`${WORKER_URL}/check/${encodeURIComponent(id)}/${encodeURIComponent(ctx)}`);
                const data = await response.json();

                if (currentDocId !== docIdAtStart) return;

                if (data.exists) {
                    uploadedFiles[ctx] = {
                        name: data.filename,
                        url: `${WORKER_URL}/view/${encodeURIComponent(id)}/${encodeURIComponent(ctx)}/${encodeURIComponent(data.filename)}`
                    };

                    // Update UI indicators
                    updateUIIndicator(ctx, true);
                }
            } catch (err) {
                console.error(`Error checking ${ctx}:`, err);
            }
        }
    }

    function updateUIIndicator(type, exists) {
        if (!exists) return;

        // Submenu items
        const idMap = {
            'LI': 'item-LI',
            'Kronologi': 'item-Kronologi',
            'PUL BAKET': 'item-PUL_BAKET'
        };

        if (idMap[type]) {
            const el = document.getElementById(idMap[type]);
            if (el) el.classList.add('has-file');
        }

        // Action boxes
        if (type === 'Penolakan') {
            if (boxPenolakan) boxPenolakan.classList.add('file-available');
        }
        if (type === 'Pemusnahan') {
            if (boxPemusnahan) {
                boxPemusnahan.classList.remove('red-box');
                boxPemusnahan.classList.add('blue-box');
                boxPemusnahan.classList.add('file-available');
            }
        }

        // If this is the context currently being viewed, show the file!
        if (type === currentUploadContext && uploadedFiles[type]) {
            showPdf(uploadedFiles[type].name, uploadedFiles[type].url);
        }
    }

    closeModalBtn.addEventListener('click', () => {
        modal.style.display = 'none';
        localStorage.removeItem('activeDocId');
        localStorage.removeItem('activeContext');
    });

    window.onclick = (event) => {
        if (event.target === modal) {
            modal.style.display = 'none';
            localStorage.removeItem('activeDocId');
            localStorage.removeItem('activeContext');
        }
    };

    function resetView() {
        // Reset preview area
        previewContainer.classList.remove('hidden');
        fileViewer.classList.add('hidden');

        // Reset uploaded files state
        uploadedFiles = {};

        // Reset UI indicators
        document.querySelectorAll('.has-file').forEach(el => el.classList.remove('has-file'));
        document.querySelectorAll('.file-available').forEach(el => el.classList.remove('file-available'));

        // Reset Pemusnahan box color
        boxPemusnahan.classList.remove('blue-box');
        boxPemusnahan.classList.add('red-box');
    }

    // Sidebar Interactions
    window.toggleSubmenu = (id) => {
        const submenu = document.getElementById(id);
        submenu.style.display = submenu.style.display === 'block' ? 'none' : 'block';
    };

    window.loadPreview = (type) => {
        currentUploadContext = type;
        localStorage.setItem('activeContext', type);

        // Check if file already exists for this context
        if (uploadedFiles[type]) {
            showPdf(uploadedFiles[type].name, uploadedFiles[type].url);
        } else {
            // Reset UI for new selection (Empty State)
            previewContainer.classList.remove('hidden');
            fileViewer.classList.add('hidden');
            document.querySelector('.empty-state p').textContent = `Upload file untuk ${type}`;
        }
    };

    // File Upload Simulation
    window.triggerFileUpload = () => {
        document.getElementById('file-input').click();
    };

    document.getElementById('file-input').addEventListener('change', async (e) => {
        if (e.target.files.length > 0) {
            const file = e.target.files[0];
            const formData = new FormData();
            formData.append('file', file);

            // Change UI to "Uploading..."
            const emptyStateText = document.querySelector('.empty-state p');
            const originalText = emptyStateText.textContent;
            emptyStateText.textContent = "Uploading to server...";

            try {
                const uploadUrl = `${WORKER_URL}/upload/${encodeURIComponent(currentDocId)}/${encodeURIComponent(currentUploadContext)}/${encodeURIComponent(file.name)}`;
                const response = await fetch(uploadUrl, {
                    method: 'PUT',
                    body: file,
                    headers: {
                        'Content-Type': file.type
                    }
                });

                if (response.ok) {
                    const fileURL = `${WORKER_URL}/view/${encodeURIComponent(currentDocId)}/${encodeURIComponent(currentUploadContext)}/${encodeURIComponent(file.name)}`;
                    uploadedFiles[currentUploadContext] = {
                        name: file.name,
                        url: fileURL
                    };

                    showPdf(file.name, fileURL);
                    updateUIIndicator(currentUploadContext, true);
                } else {
                    alert("Upload failed: " + response.statusText);
                    emptyStateText.textContent = originalText;
                }
            } catch (err) {
                console.error("Upload error:", err);
                alert("Upload error. Check console.");
                emptyStateText.textContent = originalText;
            }
        }
    });

    function showPdf(filename, fileUrl) {
        previewContainer.classList.add('hidden');
        fileViewer.classList.remove('hidden');
        viewerFilename.textContent = filename;
        pdfFrame.src = fileUrl;
    }

    window.resetUpload = () => {
        // Remove file from state
        if (currentUploadContext && uploadedFiles[currentUploadContext]) {
            delete uploadedFiles[currentUploadContext];
        }

        previewContainer.classList.remove('hidden');
        fileViewer.classList.add('hidden');
        document.getElementById('file-input').value = '';

        if (currentUploadContext === 'Pemusnahan') {
            boxPemusnahan.classList.remove('blue-box');
            boxPemusnahan.classList.add('red-box');
        }
    };

    // Theme Toggle
    const themeToggleBtn = document.getElementById('theme-toggle');
    themeToggleBtn.addEventListener('click', () => {
        document.body.classList.toggle('dark-mode');
        const icon = themeToggleBtn.querySelector('i');
        if (document.body.classList.contains('dark-mode')) {
            icon.classList.remove('fa-moon');
            icon.classList.add('fa-sun');
        } else {
            icon.classList.remove('fa-sun');
            icon.classList.add('fa-moon');
        }
    });

    // Profile Logout
    document.getElementById('profile-btn').addEventListener('click', () => {
        const dropdown = document.getElementById('profile-dropdown');
        dropdown.style.display = dropdown.style.display === 'block' ? 'none' : 'block';
    });

    document.getElementById('logout-btn').addEventListener('click', (e) => {
        e.preventDefault();
        alert('Logging out...');
        // window.location.href = 'login.html';
    });

    // Filter Logic
    window.filterTable = () => {
        const karantina = document.getElementById('filter-karantina').value;
        const permohonan = document.getElementById('filter-permohonan').value;
        const tglAwal = document.getElementById('filter-tgl-awal').value;
        const tglAkhir = document.getElementById('filter-tgl-akhir').value;

        const filteredData = tableData.filter(row => {
            const matchKarantina = karantina === '' || row.category === karantina;
            const matchPermohonan = permohonan === '' || row.permohonan === permohonan;

            let matchDate = true;
            if (tglAwal && tglAkhir) {
                matchDate = row.tgl >= tglAwal && row.tgl <= tglAkhir;
            } else if (tglAwal) {
                matchDate = row.tgl >= tglAwal;
            } else if (tglAkhir) {
                matchDate = row.tgl <= tglAkhir;
            }

            return matchKarantina && matchPermohonan && matchDate;
        });

        renderTable(filteredData);
    };

    window.exportXLS = () => {
        // Re-filter to get current view data
        const karantinaVal = document.getElementById('filter-karantina').value;
        const permohonanVal = document.getElementById('filter-permohonan').value;
        const tglAwal = document.getElementById('filter-tgl-awal').value;
        const tglAkhir = document.getElementById('filter-tgl-akhir').value;

        const filteredData = tableData.filter(row => {
            const matchKarantina = karantinaVal === '' || row.category === karantinaVal;
            const matchPermohonan = permohonanVal === '' || row.permohonan === permohonanVal;
            let matchDate = true;
            if (tglAwal && tglAkhir) { matchDate = row.tgl >= tglAwal && row.tgl <= tglAkhir; }
            else if (tglAwal) { matchDate = row.tgl >= tglAwal; }
            else if (tglAkhir) { matchDate = row.tgl <= tglAkhir; }
            return matchKarantina && matchPermohonan && matchDate;
        });

        // Header Information Helpers
        const karantinaMap = { 'Hewan': 'KH', 'Ikan': 'KI', 'Tumbuhan': 'KT' };
        const permohonanMap = {
            'Ekspor': 'EK',
            'Impor': 'IM',
            'Domestik Keluar': 'DK',
            'Domestik Masuk': 'DM'
        };

        const karantinaText = karantinaMap[karantinaVal] || 'SEMUA';
        const permohonanText = permohonanMap[permohonanVal] || 'SEMUA';
        const tglText = (tglAwal && tglAkhir) ? `${tglAwal} S/D ${tglAkhir}` : (tglAwal ? `MULAI ${tglAwal}` : (tglAkhir ? `HINGGA ${tglAkhir}` : 'SEMUA PERIODE'));
        const now = new Date();
        const waktuCetak = now.toISOString().slice(0, 19).replace('T', ' ');

        // Generate HTML for Excel with Header and 18 Columns
        let tableHTML = `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    body { font-family: Arial, sans-serif; }
                    table { border-collapse: collapse; width: 100%; margin-top: 20px; }
                    th, td { border: 1px solid #000; padding: 5px; text-align: left; }
                    center { text-align: center; }
                    h4 { margin: 5px 0; }
                </style>
            </head>
            <body>
                <center>
                    <h4>LAPORAN PENAHANAN ${karantinaText}</h4>
                    <h4>PERMOHONAN ${permohonanText}</h4>
                    <h4>7100</h4>
                    <h4>PERIODE ${tglText}</h4>
                    <h4>Waktu Cetak ${waktuCetak} &nbsp; Oleh DASHBOARD MONITORING</h4>
                    <br><br>
                </center>
                <table>
                    <thead>
                        <tr>
                            <th>No. Dokumen</th>
                            <th>Tgl Dokumen</th>
                            <th>No. P5</th>
                            <th>Tgl P5</th>
                            <th>Satpel</th>
                            <th>Pengirim</th>
                            <th>Penerima</th>
                            <th>Asal</th>
                            <th>Tujuan</th>
                            <th>Alasan</th>
                            <th>Rekomendasi</th>
                            <th>Petugas</th>
                            <th>Komoditas</th>
                            <th>Nama Tercetak</th>
                            <th>Kode HS</th>
                            <th>volume P0</th>
                            <th>volume P5</th>
                            <th>satuan</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        filteredData.forEach(row => {
            tableHTML += `
                <tr>
                    <td>${row.noDok}</td>
                    <td>${row.tgl}</td>
                    <td>-</td>
                    <td>-</td>
                    <td>${row.tkp}</td>
                    <td>${row.pengirim}</td>
                    <td>${row.penerima}</td>
                    <td>${row.asal}</td>
                    <td>${row.tujuan}</td>
                    <td>Media pembawa tidak dilaporkan kepada pejabat karantina...</td>
                    <td>dilakukan penolakan</td>
                    <td>SISTEM MONITORING</td>
                    <td>${row.komoditas}</td>
                    <td>${row.komoditas}</td>
                    <td>00000000</td>
                    <td>0,00</td>
                    <td>0,00</td>
                    <td>kemasan</td>
                </tr>
            `;
        });

        tableHTML += `
                    </tbody>
                </table>
            </body>
            </html>
        `;

        // Create Blob and Download
        const blob = new Blob([tableHTML], { type: 'application/vnd.ms-excel' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Laporan_Penahanan_${karantinaText}_${permohonanText}.xls`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    };

    // Initial Render
    filterTable();

    // Restore State from localStorage
    const savedDocId = localStorage.getItem('activeDocId');
    const savedContext = localStorage.getItem('activeContext');

    if (savedDocId) {
        openDetail(parseInt(savedDocId));
        if (savedContext) {
            // Give a small delay for checkExistingFiles to potentially fetch data
            setTimeout(() => {
                loadPreview(savedContext);
            }, 500);
        }
    }
});
