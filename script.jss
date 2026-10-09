// GRURU MUSEUM — Future Archaeology Interactive Engine
document.addEventListener('DOMContentLoaded', () => {
  initDataBooming();
  initCategoryFilter();
});

// 1. Data Booming Radial Engine
function initDataBooming() {
  const canvas = document.getElementById('boomingCanvas');
  const coreNode = document.getElementById('coreNode');
  const triggerBtn = document.getElementById('triggerBoomBtn');
  const statusDisplay = document.getElementById('boomStatus');
  const nodeCountDisplay = document.getElementById('nodeCount');
  const speedDisplay = document.getElementById('renderSpeed');
  const inspectorPanel = document.getElementById('nodeInspector');

  if (!canvas || !coreNode) return;

  const dataset = [
    { id: "GRM-TH-01", title: "แพทย์แผนไทย", cat: "วัฒนธรรม", color: "#FFC233", angle: 0, dist: 160, desc: "คัมภีร์เวชศาสตร์และสมุนไพรพื้นถิ่นเพื่อการเยียวยาองค์รวม" },
    { id: "GRM-TH-02", title: "พฤกษเคมี", cat: "ธรรมชาติ", color: "#34E89E", angle: 45, dist: 170, desc: "สารสกัดชีวภาพจากพืชเขตร้อนและงานวิจัยระดับโมเลกุล" },
    { id: "GRM-TH-03", title: "เภสัชวิเคราะห์", cat: "วิทยาศาสตร์", color: "#2EE6FF", angle: 90, dist: 155, desc: "การทดสอบสารออกฤทธิ์ทางคลินิกและการพิสูจน์เชิงประจักษ์" },
    { id: "GRM-TH-04", title: "ตำราสมุดไทย", cat: "ประวัติศาสตร์", color: "#FF5A1F", angle: 135, dist: 175, desc: "จดหมายเหตุโบราณคดีและบันทึกสูตรยาตั้งแต่สมัยกรุงศรีอยุธยา" },
    { id: "GRM-TH-05", title: "ชีวนวัตกรรม AI", cat: "เทคโนโลยี", color: "#8C6BFF", angle: 180, dist: 160, desc: "การจำลองโครงสร้างสารด้วย AI Knowledge Graph แห่งอนาคต" },
    { id: "GRM-TH-06", title: "ลายประจำยาม", cat: "ศิลปะ", color: "#FF5FA2", angle: 225, dist: 170, desc: "ลวดลายเรขาคณิตศักดิ์สิทธิ์ที่พัฒนาสู่สัญลักษณ์บรรจุภัณฑ์สากล" },
    { id: "GRM-TH-07", title: "อาหารเป็นยา", cat: "ชีวิต", color: "#34E89E", angle: 270, dist: 155, desc: "โภชนาการตามธาตุทั้งสี่ที่นำมาประยุกต์ใช้ในมื้ออาหารประจำวัน" },
    { id: "GRM-TH-08", title: "เศรษฐกิจสุขภาพ", cat: "นวัตกรรม", color: "#D7FF3A", angle: 315, dist: 175, desc: "การสร้างมูลค่าเพิ่มสู่เศรษฐกิจสร้างสรรค์ระดับโลก (Wellness Export)" }
  ];

  function runBooming() {
    // Clear previous elements
    document.querySelectorAll('.booming-node, .node-link-line').forEach(el => el.remove());

    if (statusDisplay) {
      statusDisplay.textContent = "กำลังระเบิดความรู้รอบทิศ 360°...";
      statusDisplay.style.color = "var(--gr-violet)";
    }

    const rect = canvas.getBoundingClientRect();
    const cx = canvas.clientWidth / 2;
    const cy = canvas.clientHeight / 2;

    let activeCount = 0;

    dataset.forEach((item, index) => {
      setTimeout(() => {
        const rad = (item.angle * Math.PI) / 180;
        const targetX = cx + Math.cos(rad) * item.dist;
        const targetY = cy + Math.sin(rad) * item.dist;

        // Connecting Line
        const line = document.createElement('div');
        line.className = 'node-link-line';
        line.style.left = `${cx}px`;
        line.style.top = `${cy}px`;
        line.style.width = `${item.dist}px`;
        line.style.transform = `rotate(${item.angle}deg)`;
        line.style.background = item.color;
        line.style.opacity = '0.55';
        canvas.appendChild(line);

        // Child Node
        const node = document.createElement('div');
        node.className = 'booming-node';
        node.style.left = `${targetX - 55}px`;
        node.style.top = `${targetY - 18}px`;
        node.style.borderLeft = `4px solid ${item.color}`;
        node.innerHTML = `<span>${item.title}</span><span style="font-size:0.65rem; color:#888;">[${item.cat}]</span>`;

        // Interactive Click
        node.addEventListener('click', () => {
          showNodeDetails(item);
        });

        canvas.appendChild(node);
        activeCount++;

        if (nodeCountDisplay) {
          nodeCountDisplay.textContent = `${activeCount} โหนด`;
        }

        if (index === dataset.length - 1) {
          if (statusDisplay) {
            statusDisplay.textContent = "ระเบิดสำเร็จ (เข้าใจครบใน 3 วินาที)";
            statusDisplay.style.color = "#059669";
          }
        }
      }, index * 40); // 40ms stagger from Project Brief Step 10
    });
  }

  function showNodeDetails(item) {
    if (!inspectorPanel) return;
    inspectorPanel.innerHTML = `
      <div style="font-family: var(--font-mono); font-size: 0.75rem; color: ${item.color}; font-weight: 700;">SPECIMEN // ${item.id}</div>
      <div style="font-family: var(--font-display); font-size: 1.25rem; font-weight: 700; color: var(--gr-void); margin: 0.25rem 0;">${item.title}</div>
      <div style="display: inline-block; background: ${item.color}22; color: #111; font-size: 0.75rem; padding: 0.2rem 0.5rem; border-radius: 4px; font-weight: 600; margin-bottom: 0.5rem;">หมวด: ${item.cat}</div>
      <p style="font-size: 0.9rem; color: var(--gr-text-secondary); line-height: 1.6;">${item.desc}</p>
      <div style="margin-top: 0.85rem; padding-top: 0.85rem; border-top: 1px dashed var(--gr-grid-light); font-size: 0.8rem; color: var(--gr-fog);">
        ความสัมพันธ์: <strong>influenced by</strong> [ภูมิปัญญาไทย] • แหล่งที่มา: คลังจดหมายเหตุและตำราแพทย์หลวง
      </div>
    `;
  }

  if (triggerBtn) triggerBtn.addEventListener('click', runBooming);
  if (coreNode) coreNode.addEventListener('click', runBooming);

  // Auto trigger after 1 second for engaging demonstration
  setTimeout(runBooming, 800);
}

// 2. Category Interactive Filter
function initCategoryFilter() {
  const cards = document.querySelectorAll('.cat-box');
  cards.forEach(card => {
    card.addEventListener('mouseenter', () => {
      card.style.borderColor = 'var(--gr-lime)';
      card.style.boxShadow = '0 8px 20px rgba(0,0,0,0.06)';
    });
    card.addEventListener('mouseleave', () => {
      card.style.borderColor = 'var(--gr-grid-light)';
      card.style.boxShadow = 'none';
    });
  });
}
