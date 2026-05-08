// Header scroll effect
const header = document.getElementById('site-header');
if (header) {
  window.addEventListener('scroll', () => {
    header.classList.toggle('scrolled', window.scrollY > 40);
  }, { passive: true });
}

// Hamburger menu
const hamburger = document.getElementById('hamburger');
const nav = document.getElementById('global-nav');
if (hamburger && nav) {
  hamburger.addEventListener('click', () => {
    hamburger.classList.toggle('open');
    nav.classList.toggle('open');
  });
  nav.querySelectorAll('a').forEach(a => {
    a.addEventListener('click', () => {
      hamburger.classList.remove('open');
      nav.classList.remove('open');
    });
  });
}

// FAQ accordion
document.querySelectorAll('.faq-question').forEach(btn => {
  btn.addEventListener('click', () => {
    const item = btn.closest('.faq-item');
    const isOpen = item.classList.contains('open');
    document.querySelectorAll('.faq-item.open').forEach(i => i.classList.remove('open'));
    if (!isOpen) item.classList.add('open');
  });
});

// Contact form
const form = document.getElementById('contact-form');
if (form) {
  form.addEventListener('submit', async e => {
    e.preventDefault();
    if (!form.checkValidity()) {
      form.reportValidity();
      return;
    }

    const btn = form.querySelector('.form-submit');
    const originalText = btn.textContent;
    btn.disabled = true;
    btn.textContent = '送信中...';

    try {
      const data = new FormData(form);
      await fetch('/', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams(data).toString()
      });
      form.style.display = 'none';
      const success = document.getElementById('form-success');
      if (success) success.style.display = 'block';
    } catch (error) {
      btn.disabled = false;
      btn.textContent = originalText;
      alert('送信に失敗しました。恐れ入りますが、LINEからお問い合わせください。');
    }
  });
}

// Smooth appear on scroll
const observer = new IntersectionObserver(entries => {
  entries.forEach(e => { if (e.isIntersecting) { e.target.classList.add('visible'); } });
}, { threshold: 0.1 });
document.querySelectorAll('.service-card, .case-card, .about-card, .faq-item, .starter-panel, .guidebot-panel').forEach(el => {
  el.style.opacity = '0';
  el.style.transform = 'translateY(20px)';
  el.style.transition = 'opacity .5s ease, transform .5s ease';
  observer.observe(el);
});
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('.visible').forEach(el => {
    el.style.opacity = '1';
    el.style.transform = 'none';
  });
});
// Re-trigger for already visible
setTimeout(() => {
  document.querySelectorAll('[style*="opacity: 0"]').forEach(el => {
    const rect = el.getBoundingClientRect();
    if (rect.top < window.innerHeight) {
      el.style.opacity = '1';
      el.style.transform = 'none';
    }
  });
}, 100);

// API-free guide bot
const guideBotResult = document.getElementById('guidebot-result');
const guideBotOptions = document.querySelectorAll('.guidebot-option');
const guideBotAnswers = {
  files: {
    title: 'おすすめ：PC・クラウド整理スターター',
    text: '保存場所を棚卸しして、デスクトップ・Downloads・OneDrive・iCloudの役割を分けます。まずは「どこに何があるか分かる状態」を作ります。'
  },
  customers: {
    title: 'おすすめ：顧客管理フォルダ設計',
    text: '顧客別フォルダ、案件メモ、提案書、契約書、請求書の置き場を作ります。初回成約後の管理にもつながる形です。'
  },
  documents: {
    title: 'おすすめ：請求・提案資料の整理',
    text: '見積、請求、提案、規約などを分け、探す時間を減らします。必要ならテンプレート化や月額サポートにも進めます。'
  },
  automation: {
    title: 'おすすめ：情報整理からAI活用へ',
    text: 'いきなり自動化せず、まずデータと書類の置き場を整えます。その後、AIに任せてよい作業と人が見る作業を分けます。'
  }
};
if (guideBotResult && guideBotOptions.length) {
  guideBotOptions.forEach(button => {
    button.addEventListener('click', () => {
      guideBotOptions.forEach(item => item.classList.remove('is-active'));
      button.classList.add('is-active');
      const answer = guideBotAnswers[button.dataset.guide];
      if (!answer) return;
      guideBotResult.innerHTML = `<p class="guidebot-result-title">${answer.title}</p><p class="guidebot-result-text">${answer.text}</p>`;
    });
  });
}
