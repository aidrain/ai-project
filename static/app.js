let currentState = null;

function setBanner(message, type = "info") {
  const banner = document.getElementById("status-banner");
  banner.textContent = message;
  banner.className = `banner ${type}`;
}

function escapeHtml(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function fillCategorySelectors(courses) {
  const defaultCats = ["전공핵심", "전공선택", "전공기초교양", "진로취업"];
  const foundCats = courses.map((c) => c["과목구분"]).filter(Boolean);
  const allCats = [...new Set([...defaultCats, ...foundCats])];

  ["add-cat", "edit-cat"].forEach((id) => {
    const sel = document.getElementById(id);
    const prev = sel.value;
    sel.innerHTML = allCats
      .map((cat) => `<option value="${escapeHtml(cat)}">${escapeHtml(cat)}</option>`)
      .join("");
    if (prev && allCats.includes(prev)) {
      sel.value = prev;
    } else if (id === "add-cat" && allCats.includes("전공선택")) {
      sel.value = "전공선택";
    }
  });
}

function fillCourseSelectors(courses) {
  fillCategorySelectors(courses);
  const editSel = document.getElementById("edit-target");
  const delSel = document.getElementById("del-target");
  const prevEdit = editSel.value;
  const prevDel = delSel.value;

  const optionsHtml = courses
    .map((c) => `<option value="${escapeHtml(c["교과목명"])}">[${escapeHtml(c["과목구분"])}] ${escapeHtml(c["교과목명"])} (${escapeHtml(c["학년-학기"])})</option>`)
    .join("");

  editSel.innerHTML = optionsHtml;
  delSel.innerHTML = optionsHtml;

  if (prevEdit && courses.some((c) => c["교과목명"] === prevEdit)) {
    editSel.value = prevEdit;
  }
  if (prevDel && courses.some((c) => c["교과목명"] === prevDel)) {
    delSel.value = prevDel;
  }
  syncEditFormWithSelectedCourse();
}

function syncEditFormWithSelectedCourse() {
  if (!currentState || !currentState.courses_2027) return;
  const target = document.getElementById("edit-target").value;
  const course = currentState.courses_2027.find((c) => c["교과목명"] === target);
  if (!course) return;

  document.getElementById("edit-cat").value = course["과목구분"];
  document.getElementById("edit-name").value = course["교과목명"];
  document.getElementById("edit-term").value = course["학년-학기"];
  const [cr = 3, th = 3, pr = 0] = String(course["학점-이론-실습"]).split("-").map(Number);
  document.getElementById("edit-credit").value = cr;
  document.getElementById("edit-theory").value = th;
  document.getElementById("edit-practice").value = pr;
}

function renderState(data, statusMsg = null) {
  currentState = data;
  if (!data.loaded) {
    setBanner("2026학년도 교육과정 엑셀(.xlsx) 파일을 업로드해 주세요.", "info");
    return;
  }

  if (statusMsg) {
    setBanner(statusMsg, "success");
  } else {
    setBanner(`불러온 파일: ${data.filename} (${data.meta.college} · ${data.meta.major})`, "info");
  }

  const m = data.metrics;
  document.getElementById("m-count-2026").textContent = `${m.count_2026}개`;
  document.getElementById("m-sub-2026").innerHTML = (m.categories_2026 || [])
    .map((c) => `${escapeHtml(c.name)} (${c.count}과목 / 소계: ${escapeHtml(c.subtotal)})`)
    .join("<br />");
  document.getElementById("m-count-2027").textContent = `${m.count_2027}개`;
  document.getElementById("m-delta-2027").textContent = m.delta_text;
  document.getElementById("m-credits-2026").textContent = m.credits_2026;
  document.getElementById("m-credits-2027").textContent = m.credits_2027;

  fillCourseSelectors(data.courses_2027);

  // 교과목 개정표 렌더링
  const revTable = document.getElementById("revision-table");
  const revSummary = document.getElementById("revision-summary");
  const undoBtn = document.getElementById("btn-undo");

  if (data.revision_rows.length === 0) {
    revTable.style.display = "none";
    undoBtn.style.display = "none";
    revSummary.textContent = "아직 추가·변경·폐지한 내역이 없습니다.";
  } else {
    revTable.style.display = "table";
    undoBtn.style.display = "inline-block";
    revSummary.textContent = `총 ${data.revision_rows.length}건 (신설 ${m.add_count}건 · 변경 ${m.edit_count}건 · 폐지 ${m.delete_count}건)`;
    revTable.querySelector("tbody").innerHTML = data.revision_rows
      .map(
        (r) => `
        <tr>
          <td><span class="badge badge-${escapeHtml(r["구분"])}">${escapeHtml(r["구분"])}</span></td>
          <td>${escapeHtml(r["과목구분"])}</td>
          <td class="left">${escapeHtml(r["변경 전 (2026)"])}</td>
          <td class="left">${escapeHtml(r["변경 후 (2027)"])}</td>
        </tr>`
      )
      .join("");
  }

  // 2027학년도 편성표 렌더링
  const container = document.getElementById("curriculum-tables");
  container.innerHTML = data.categories
    .map(
      (cat) => `
      <div class="cat-title">${escapeHtml(cat.name)} (${cat.count}과목 / 소계: ${escapeHtml(cat.subtotal)})</div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>과목구분</th>
              <th>학수번호</th>
              <th>교과목명</th>
              <th>학년-학기</th>
              <th>학점-이론-실습</th>
            </tr>
          </thead>
          <tbody>
            ${cat.rows
              .map(
                (row) => `
                <tr>
                  <td>${escapeHtml(row["과목구분"])}</td>
                  <td>${escapeHtml(row["학수번호"] || "-")}</td>
                  <td class="left">${escapeHtml(row["교과목명"])}</td>
                  <td>${escapeHtml(row["학년-학기"])}</td>
                  <td>${escapeHtml(row["학점-이론-실습"])}</td>
                </tr>`
              )
              .join("")}
          </tbody>
        </table>
      </div>`
    )
    .join("");
}

async function callApi(url, options = {}, successMsg = null) {
  console.log("[요청]", url);
  const res = await fetch(url, options);
  const data = await res.json();
  if (!res.ok) {
    setBanner(data.error || "요청 처리 중 오류가 발생했습니다.", "error");
    return;
  }
  renderState(data, successMsg);
}

// 탭 전환
document.querySelectorAll(".tab-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
    document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById(btn.dataset.tab).classList.add("active");
  });
});

document.getElementById("edit-target").addEventListener("change", syncEditFormWithSelectedCourse);

document.getElementById("btn-upload").addEventListener("click", async () => {
  const fileInput = document.getElementById("excel-file");
  if (!fileInput.files.length) {
    setBanner("업로드할 .xlsx 파일을 먼저 선택해 주세요.", "error");
    return;
  }
  const formData = new FormData();
  formData.append("file", fileInput.files[0]);
  await callApi("/api/upload", { method: "POST", body: formData }, "새 엑셀 파일을 불러왔습니다.");
});

document.getElementById("form-add").addEventListener("submit", async (e) => {
  e.preventDefault();
  const payload = {
    과목구분: document.getElementById("add-cat").value,
    학수번호: "",
    교과목명: document.getElementById("add-name").value,
    "학년-학기": document.getElementById("add-term").value,
    학점: Number(document.getElementById("add-credit").value),
    이론: Number(document.getElementById("add-theory").value),
    실습: Number(document.getElementById("add-practice").value),
  };
  await callApi(
    "/api/change/add",
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    },
    "신규 교과목 개설이 반영되었습니다."
  );
  document.getElementById("add-name").value = "";
});

document.getElementById("form-edit").addEventListener("submit", async (e) => {
  e.preventDefault();
  const payload = {
    대상과목명: document.getElementById("edit-target").value,
    새_과목구분: document.getElementById("edit-cat").value,
    새_교과목명: document.getElementById("edit-name").value,
    새_학년학기: document.getElementById("edit-term").value,
    학점: Number(document.getElementById("edit-credit").value),
    이론: Number(document.getElementById("edit-theory").value),
    실습: Number(document.getElementById("edit-practice").value),
  };
  await callApi(
    "/api/change/edit",
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    },
    "교과목 변경이 반영되었습니다."
  );
});

document.getElementById("form-del").addEventListener("submit", async (e) => {
  e.preventDefault();
  const payload = { 대상과목명: document.getElementById("del-target").value };
  await callApi(
    "/api/change/delete",
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    },
    "교과목 폐설이 반영되었습니다."
  );
});

document.getElementById("btn-undo").addEventListener("click", async () => {
  await callApi("/api/change/undo", { method: "POST" }, "마지막 개정 작업을 취소했습니다.");
});

document.getElementById("btn-reset").addEventListener("click", async () => {
  await callApi("/api/change/reset", { method: "POST" }, "개정 작업을 모두 초기화했습니다.");
});

callApi("/api/state");
