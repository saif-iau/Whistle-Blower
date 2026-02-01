import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './Ssamic2WebPart.module.scss';
import * as strings from 'Ssamic2WebPartStrings';

export interface ISsamic2WebPartProps {
  description: string;
  apiEndpoint: string;
}

interface AppState {
  language: 'en' | 'ar';
  mobileMenuOpen: boolean;
  formData: {
    name: string;
    email: string;
    site: string;
    department: string;
    message: string;
    file: File | null;
  };
  submitting: boolean;
  errors: Record<string, string>;
  touched: Set<string>;
}

export default class Ssamic2WebPart extends BaseClientSideWebPart<ISsamic2WebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _state: AppState;
  private _container: HTMLElement | null = null;
  private readonly MAX_FILE_SIZE = 10 * 1024 * 1024; // 10 MB in bytes

  constructor() {
    super();
    this._state = {
      language: 'en',
      mobileMenuOpen: false,
      formData: { name: '', email: '', site: '', department: '', message: '', file: null },
      submitting: false,
      errors: {},
      touched: new Set()
    };
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div id="whistleblower-app"></div>
      <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/toastify-js/src/toastify.min.css">
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');
        
        #whistleblower-app * { 
          margin: 0; 
          padding: 0; 
          box-sizing: border-box; 
          font-family: 'Roboto', sans-serif;
        }
        
        [dir="rtl"] { direction: rtl; }
        details summary { list-style: none; cursor: pointer; }
        details summary::-webkit-details-marker { display: none; }
        
        .nav-btn { 
          background: transparent; 
          color: #F2F2F2; 
          border: none; 
          padding: 8px 16px; 
          cursor: pointer; 
          font-size: 1rem;
          transition: all 0.3s ease;
        }
        .nav-btn:hover { 
          background: rgba(242, 178, 155, 0.2); 
          border-radius: 4px;
          color: #F2B29B;
        }
        
        @media (min-width: 768px) {
          .mobile-only { display: none !important; }
          .desktop-nav { display: flex !important; }
        }
        @media (max-width: 767px) {
          .mobile-only { display: block !important; }
          .desktop-nav { display: none !important; }
        }

        
.submit-btn {
  width: 100%;
  padding: 20px 16px; /* increased vertical padding for greater height */
  min-height: 56px; /* ensure a taller button across layouts */
  background-color: #0078D4; /* Fluent UI Blue */
  color: #FFFFFF;
  border: none;
  border-radius: 8px;
  font-size: 1.125rem;
  font-weight: 600;
  cursor: pointer;
  transition: background-color 0.2s ease, transform 0.05s ease, box-shadow 0.2s ease, opacity 0.2s ease;
  box-shadow: 0 4px 6px rgba(0,0,0,0.1);
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 10px;
}

/* Hover (only when not disabled) */
.submit-btn:not(:disabled):hover {
  background-color: #005A9E; /* Darker hover */
}

/* Active (click press) */
.submit-btn:not(:disabled):active {
  transform: translateY(1px);
  box-shadow: 0 3px 5px rgba(0,0,0,0.12);
}

/* Focus-visible for keyboard users */
.submit-btn:focus-visible {
  outline: none;
  box-shadow: 0 0 0 3px rgba(0,120,212,0.35);
}

/* Disabled state */
.submit-btn:disabled {
  cursor: not-allowed;
  opacity: 0.6;
}

/* Content wrapper */
.btn-content {
  display: inline-flex;
  align-items: center;
  gap: 10px;
}

/* Spinner */
.spinner {
  width: 18px;
  height: 18px;
  border: 2px solid rgba(255,255,255,0.5);
  border-top-color: #FFFFFF;
  border-radius: 50%;
  display: inline-block;
  animation: spin 0.8s linear infinite;
}

@keyframes spin {
  to { transform: rotate(360deg); }
}

/* Optional: dark-mode override */
@media (prefers-color-scheme: dark) {
  .submit-btn {
    box-shadow: 0 4px 6px rgba(0,0,0,0.35);
  }
}

      </style>
    `;

    void this._loadToastify().then(() => {
      this._container = this.domElement.querySelector('#whistleblower-app');
      this._render();
      this._attachEventListeners();
    });
  }

  private _loadToastify(): Promise<void> {
    return new Promise((resolve) => {
      if (typeof (window as any).Toastify !== 'undefined') {
        resolve();
        return;
      }

      const script = document.createElement('script');
      script.src = 'https://cdn.jsdelivr.net/npm/toastify-js';
      script.onload = () => resolve();
      document.head.appendChild(script);
    });
  }

  private _showToast(message: string, type: 'success' | 'error' | 'warning' = 'success'): void {
    const bgColor = type === 'success' ? '#10b981' :  // Green
      type === 'error' ? '#ef4444' :  // Red
        '#f97316';  // Orange

    (window as any).Toastify({
      text: message,
      duration: 4000,
      gravity: "top",
      position: "center",
      style: { background: bgColor }
    }).showToast();
  }

  private _getTranslations() {
    return {
      en: {
        nav: { home: "Home", form: "Submit Report", faq: "FAQ", committee: "Committee" },
        hero: {
          title1: "Speak Up Safely",
          title2: "Make a Difference",
          subtitle: "Your voice matters. Report concerns anonymously and securely.",
          ctaReport: "Submit Report",
          ctaCommittee: "Learn About the Committee"
        },
        form: {
          title: "Submit Your Report",
          subtitle: "All information is treated confidentially",
          name: "Name (Optional)",
          namePlaceholder: "Enter your name or leave blank",
          email: "Email (Optional)",
          emailPlaceholder: "your.email@example.com",
          location: "Location *",
          department: "Management (Optional)",
          message: "Report Details *",
          messagePlaceholder: "Describe the incident or concern in detail...",
          submitBtn: "Submit Report Securely",
          submitting: "Submitting...",
          select: "-- Select --",

          file: "Attach File (Optional)",
          fileHelp: "Max size: 10 MB. All types accepted.",


        },
        faq: {
          title: "Frequently Asked Questions",
          questions: [
            { q: "What is the Culture & Workplace Integrity Committee?", a: "A permanent committee established at Alkhorayef Industries – Military Sector responsible for reviewing employees' behavioral complaints and observations." },
            { q: "What are the committee's main roles?", a: "Studying all complaints and observations received from employees; verifying the validity of allegations and facts by listening to the concerned parties and gathering necessary data and documents; investigating cases and submitting recommendations." },
            { q: "How does the committee handle conflicts of interest?", a: "If a complaint concerns a committee member or their department, the concerned member will be temporarily excluded and another member may be added as a replacement." },
            { q: "Who can submit a report?", a: "Any employee, contractor, or third party who wishes to raise workplace concerns can submit a report." },
            { q: "Can I attach evidence to my report?", a: "Yes. You may attach files (maximum 10 MB) to support your report." },
            { q: "Will I face retaliation for reporting?", a: "No. There is a strict no-retaliation policy and investigations are handled confidentially to protect reporters." }
          ]
        },
        footer: "© 2026 AMIC IT. All rights reserved. Confidential & Secure Reporting.",
        errors: {
          invalidEmail: "Invalid email",
          required: "Required",
          selectLocation: "Select location",
          fileTooLarge: "File exceeds 10 MB limit",
        },
        success: "Report submitted successfully!",
        errorSubmit: "Error submitting report"
      },
      committee: {
        title: 'Administrative Decision',
        heading: 'Culture & Workplace Integrity Committee',
        date: '28.10.2025',
        intro: 'Alkhorayef Industries – Military Sector is committed to promoting a healthy work culture where fairness and transparency are enforced, resulting in:',
        formationTitle: 'First: Formation of the Committee',
        formationText: "A permanent committee is established at Alkhorayef Industries – Military Sector under the name 'Culture & Workplace Integrity Committee' which will be responsible for reviewing employees' behavioral complaints and observations.",
        roleTitle: 'Second: Committee Role',
        roleItems: [
          'Studying all complaints and observations received from employees.',
          'Verifying the validity of allegations and facts by listening to the concerned parties and gathering necessary data and documents.',
          'Investigating cases and submitting recommendations.'
        ],
        conflictTitle: 'Third: Conflict of Interest',
        conflictText: 'In the event of a conflict of interest, such as a complaint concerning a committee member or their department, the concerned committee member shall be temporarily excluded, with the possibility of adding another member as a replacement.'
      },
      ar: {
        nav: { home: "الرئيسية", form: "تقديم بلاغ", faq: "الأسئلة الشائعة", committee: "اللجنة" },
        hero: {
          title1: "تحدث بأمان",
          title2: "أحدث فرقًا",
          subtitle: "صوتك مهم. قدم مخاوفك بشكل مجهول وآمن.",
          ctaReport: "تقديم بلاغ",
          ctaCommittee: "تعرف على اللجنة"
        },
        form: {
          title: "تقديم بلاغك",
          subtitle: "يتم التعامل مع جميع المعلومات بسرية",
          name: "الاسم (اختياري)",
          namePlaceholder: "أدخل اسمك أو اتركه فارغًا",
          email: "البريد الإلكتروني (اختياري)",
          emailPlaceholder: "your.email@example.com",
          location: "الموقع *",
          department: "الإدارة (اختياري)",
          message: "تفاصيل البلاغ *",
          messagePlaceholder: "صف الحادثة أو المخاوف بالتفصيل...",
          submitBtn: "تقديم البلاغ بشكل آمن",
          submitting: "جارٍ الإرسال...",
          select: "-- اختر --",

          file: "إرفاق ملف (اختياري)",
          fileHelp: "الحد الأقصى: 10 ميجابايت. جميع الأنواع مقبولة.",

        },
        faq: {
          title: "الأسئلة المتكررة",
          questions: [
            { q: "ما هي لجنة نزاهة بيئة العمل؟", a: "لجنة دائمة في شركة الخريف - القطاع العسكري مسؤولة عن مراجعة شكاوى الموظفين والملاحظات السلوكية." },
            { q: "ما هي الأدوار الرئيسية للجنة؟", a: "دراسة جميع الشكاوى والملاحظات الواردة من الموظفين؛ التحقق من صحة الادعاءات والوقائع بالاستماع إلى الأطراف المعنية وجمع البيانات والمستندات اللازمة؛ التحقيق في القضايا وتقديم التوصيات." },
            { q: "كيف تتعامل اللجنة مع حالات تعارض المصالح؟", a: "في حال تعلق الشكوى بعضو في اللجنة أو قسمه، يُستبعد العضو المعني مؤقتًا، مع إمكانية إضافة عضو آخر كبديل." },
            { q: "من يمكنه تقديم بلاغ؟", a: "يمكن لأي موظف أو متعاقد أو طرف خارجي تقديم بلاغ بشأن مخاوف بيئة العمل." },
            { q: "هل يمكنني إرفاق أدلة؟", a: "نعم. يمكنك إرفاق ملفات (بحد أقصى 10 ميجابايت) لدعم بلاغك." },
            { q: "هل سأتعرض لانتقام؟", a: "لا. هناك سياسة صارمة لعدم الانتقام ويتم التعامل مع التحقيقات بسرية لحماية المبلغين." }
          ]
        },
        footer: "© 2026 AMIC. جميع الحقوق محفوظة. إبلاغ سري وآمن.",
        errors: {
          invalidEmail: "بريد إلكتروني غير صالح",
          required: "مطلوب",
          selectLocation: "اختر موقع",

          fileTooLarge: "حجم الملف يتجاوز 10 ميجابايت",

        },
        success: "تم تقديم البلاغ بنجاح!",
        errorSubmit: "خطأ في تقديم البلاغ"
      }
    };
  }

  private _getSitesData() {
    return {
      en: [
        { id: 'Riyadh HQ', name: 'Riyadh HQ' },
        { id: 'Jeddah', name: 'Jeddah' },
        { id: 'Dammam', name: 'Dammam' },
        { id: 'Khasm Alan', name: 'Khasm Alan' },
        { id: 'Taif', name: 'Taif' },
        { id: 'Qassim', name: 'Qassim' },
        { id: 'Hofuf', name: 'Hofuf' },
        { id: 'Medina', name: 'Medina' }
      ],
      ar: [
        { id: 'Riyadh HQ', name: 'الرياض - المركز الرئيسي' },
        { id: 'Jeddah', name: 'جدة' },
        { id: 'Dammam', name: 'الدمام' },
        { id: 'Khasm Alan', name: 'خشم العان' },
        { id: 'Taif', name: 'الطائف' },
        { id: 'Qassim', name: 'القصيم' },
        { id: 'Hofuf', name: 'الهفوف' },
        { id: 'Medina', name: 'المدينة المنورة' }
      ]
    };
  }

  private _getDepartmentsData() {
    return {
      en: [
        { id: 'Business Development', name: 'Business Development' },
        { id: 'Business Process', name: 'Business Process' },
        { id: 'Contracts and Compliance', name: 'Contracts and Compliance' },
        { id: 'Engineering & RD', name: 'Engineering & RD' },
        { id: 'Executive', name: 'Executive' },
        { id: 'Finance', name: 'Finance' },
        { id: 'HR', name: 'HR' },
        { id: 'IPP', name: 'IPP' },
        { id: 'IT & MIS', name: 'IT & MIS' },
        { id: 'Projects', name: 'Projects' },
        { id: 'Supply Chain', name: 'Supply Chain' }
      ],
      ar: [
        { id: 'Business Development', name: 'تطوير الأعمال' },
        { id: 'Business Process', name: 'عمليات الأعمال' },
        { id: 'Contracts and Compliance', name: 'العقود والامتثال' },
        { id: 'Engineering & RD', name: 'الهندسة والبحث والتطوير' },
        { id: 'Executive', name: 'الإدارة التنفيذية' },
        { id: 'Finance', name: 'المالية' },
        { id: 'HR', name: 'الموارد البشرية' },
        { id: 'IPP', name: 'برنامج المشاركة الصناعية' },
        { id: 'IT & MIS', name: 'تقنية المعلومات ونظم المعلومات الإدارية' },
        { id: 'Projects', name: 'المشاريع' },
        { id: 'Supply Chain', name: 'سلسلة الإمداد' }
      ]
    };
  }


  private _validateField(
    name: string,
    value: string | File | null | undefined
  ): string | null {
    const t = this._getTranslations()[this._state.language].errors;

    // Email (optional but must be valid if provided)
    if (name === 'email') {
      const v = (value ?? '') as string;
      if (v && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) {
        return t.invalidEmail;
      }
      return null;
    }

    // Message (required)
    if (name === 'message') {
      const v = (value ?? '') as string;
      if (!v.trim()) {
        return t.required;
      }
      return null;
    }

    // Site (required - must be selected)
    if (name === 'site') {
      const v = (value ?? '') as string;
      if (!v) {
        return t.selectLocation;
      }
      return null;
    }




    if (name === 'file') {


      const file = (value instanceof File) ? value : null;

      if (file) {
        const MAX_BYTES = 10 * 1048576;
        if (file.size > MAX_BYTES) {
          return t.fileTooLarge;
        }
      }
      return null;
    }


    return null;
  }


  private _render(): void {
    if (!this._container) return;

    const t = this._getTranslations()[this._state.language];
    const sites = this._getSitesData()[this._state.language];
    const departments = this._getDepartmentsData()[this._state.language];
    const isRTL = this._state.language === 'ar';

    this._container.innerHTML = `
      <div class="min-h-screen" dir="${isRTL ? 'rtl' : 'ltr'}" style="background: linear-gradient(135deg, #293940 0%, #3E5159 100%); min-height: 100vh;">
        <div style="max-width: 1400px; margin: 0 auto; padding: 40px 20px;">

          <!-- Navigation -->
          <nav style="width: 100%; background: rgba(140, 137, 121, 0.15); border-bottom: 1px solid rgba(242, 178, 155, 0.3); border-radius: 8px; margin-bottom: 40px; backdrop-filter: blur(10px);">
            <div style="max-width: 1200px; margin: 0 auto; padding: 0 20px; display: flex; justify-content: space-between; align-items: center; min-height: 64px; direction: ltr;">
              <div style="display: flex; align-items: center; gap: 12px;">
                <img src="${require('./assets/logo-bg.png')}" alt="Logo" style="height: 50px; margin: 0; display: block; border: 1px solid rgba(242, 178, 155, 0.2); border-radius: 12px; box-shadow: 0 12px 30px rgba(41,57,64,0.35); background: #F2F2F2;" />
              </div>
              
              <div class="desktop-nav" style="display: none; gap: 20px;">
                <button class="nav-btn" data-action="scrollTo" data-target="home">${t.nav.home}</button>
                <button class="nav-btn" data-action="scrollTo" data-target="form">${t.nav.form}</button>
                <button class="nav-btn" data-action="scrollTo" data-target="faq">${t.nav.faq}</button>
              </div>

              <div style="display: flex; align-items: center; gap: 10px;">
                <button data-action="toggleLanguage" style="background: #8C8979; color: #F2F2F2; padding: 6px 12px; border-radius: 4px; border: none; cursor: pointer; transition: all 0.3s ease;">
                  ${this._state.language === 'en' ? 'العربية' : 'English'}
                </button>
                <button class="mobile-only" data-action="toggleMenu" style="display: none; color: #F2F2F2; background: transparent; border: none; font-size: 24px; cursor: pointer;">☰</button>
              </div>
            </div>

            ${this._state.mobileMenuOpen ? `
              <div class="mobile-only" style="background: rgba(140, 137, 121, 0.15); padding: 20px; text-align: center; border-top: 1px solid rgba(242, 178, 155, 0.3); direction: ltr;">
                <div style="display:flex; justify-content:center; align-items:center; margin-bottom:12px;">
                  <img src="${require('./assets/logo-bg.png')}" alt="Logo" style="height:40px; margin:0; border: 1px solid rgba(242,178,155,0.15); border-radius:10px; background:#F2F2F2;" />
                </div>
                <button class="nav-btn" data-action="scrollTo" data-target="home" style="display: block; width: 100%; text-align: center; margin: 10px 0;">${t.nav.home}</button>
                <button class="nav-btn" data-action="scrollTo" data-target="form" style="display: block; width: 100%; text-align: center; margin: 10px 0;">${t.nav.form}</button>
                <button class="nav-btn" data-action="scrollTo" data-target="faq" style="display: block; width: 100%; text-align: center; margin: 10px 0;">${t.nav.faq}</button>
              </div>
            ` : ''}
          </nav>

          <div style="max-width: 1200px; margin: 0 auto;">
            
            <!-- Hero -->
            <section id="home" style="display: flex; align-items: center; justify-content: center; color: #F2F2F2; margin-bottom: 60px;">
              <div style="width: 100%; max-width: 1200px; display:flex; flex-direction: ${isRTL ? 'row-reverse' : 'row'}; gap:24px; align-items:center; padding: 40px; border-radius: 12px; background: linear-gradient(90deg, rgba(41,57,64,0.9), rgba(62,81,89,0.7)); box-shadow: 0 20px 50px rgba(0,0,0,0.45);">
                <div style="flex:1; text-align: ${isRTL ? 'right' : 'left'};">
                  <h1 style="font-size: 3rem; font-weight: 700; margin:0 0 12px 0; color: #F2B29B; line-height:1.05; text-shadow: 2px 2px 6px rgba(0,0,0,0.4);">
                    ${t.hero.title1}<br/>${t.hero.title2}
                  </h1>
                  <p style="font-size: 1.125rem; opacity: 0.95; color: #F2F2F2; margin-bottom: 18px;">${t.hero.subtitle}</p>
                </div>

               
              </div>
            </section>

            <!-- Form -->
            <div id="form" style="background: #F2F2F2; border-radius: 16px; padding: 40px; box-shadow: 0 20px 60px rgba(41, 57, 64, 0.5); margin-bottom: 60px; border: 1px solid rgba(242, 178, 155, 0.2);">
              <h2 style="font-size: 2rem; font-weight: bold; margin-bottom: 1rem; text-align: center; color: #293940;">${t.form.title}</h2>
              <p style="text-align: center; color: #8C8979; margin-bottom: 2rem;">${t.form.subtitle}</p>
              
              <form id="reportForm">
                <div style="margin-bottom: 1.5rem;">
                  <label style="display: block; margin-bottom: 0.5rem; font-weight: 500; color: #293940;">${t.form.name}</label>
                  <input type="text" name="name" value="${this._state.formData.name}" placeholder="${t.form.namePlaceholder}" style="width: 100%; padding: 12px; border: 1px solid #8C8979; border-radius: 8px; font-size: 1rem; background: white; transition: border 0.3s ease;" />
                </div>

                <div style="margin-bottom: 1.5rem;">
                  <label style="display: block; margin-bottom: 0.5rem; font-weight: 500; color: #293940;">${t.form.email}</label>
                  <input type="email" name="email" value="${this._state.formData.email}" placeholder="${t.form.emailPlaceholder}" style="width: 100%; padding: 12px; border: 1px solid ${this._state.errors.email ? '#ef4444' : '#8C8979'}; border-radius: 8px; font-size: 1rem; background: white; transition: border 0.3s ease;" />
                  <p data-error="email" style="color: #ef4444; font-size: 0.875rem; margin-top: 0.5rem; display: ${this._state.errors.email ? 'block' : 'none'};">${this._state.errors.email || ''}</p>
                </div>

                <div style="margin-bottom: 1.5rem;">
                  <label style="display: block; margin-bottom: 0.5rem; font-weight: 500; color: #293940;">${t.form.department}</label>
                  <select name="department" style="width: 100%; padding: 12px; border: 1px solid ${this._state.errors.department ? '#ef4444' : '#8C8979'}; border-radius: 8px; font-size: 1rem; background: white; transition: border 0.3s ease;">
                    <option value="">${t.form.select}</option>
                    ${departments.map(d => `<option value="${d.id}" ${this._state.formData.department === d.id ? 'selected' : ''}>${d.name}</option>`).join('')}
                  </select>
                </div>

                <div style="margin-bottom: 1.5rem;">
                  <label style="display: block; margin-bottom: 0.5rem; font-weight: 500; color: #293940;">${t.form.location}</label>
                  <select name="site" style="width: 100%; padding: 12px; border: 1px solid ${this._state.errors.site ? '#ef4444' : '#8C8979'}; border-radius: 8px; font-size: 1rem; background: white; transition: border 0.3s ease;">
                    <option value="">${t.form.select}</option>
                    ${sites.map(s => `<option value="${s.id}" ${this._state.formData.site === s.id ? 'selected' : ''}>${s.name}</option>`).join('')}
                  </select>
                  <p data-error="site" style="color: #ef4444; font-size: 0.875rem; margin-top: 0.5rem; display: ${this._state.errors.site ? 'block' : 'none'};">${this._state.errors.site || ''}</p>
                </div>

                <div style="margin-bottom: 1.5rem;">
                  <label style="display: block; margin-bottom: 0.5rem; font-weight: 500; color: #293940;">${t.form.message}</label>
                  <textarea name="message" rows="6" placeholder="${t.form.messagePlaceholder}" style="width: 100%; padding: 12px; border: 1px solid ${this._state.errors.message ? '#ef4444' : '#8C8979'}; border-radius: 8px; font-size: 1rem; resize: vertical; background: white; transition: border 0.3s ease;">${this._state.formData.message}</textarea>
                  <p data-error="message" style="color: #ef4444; font-size: 0.875rem; margin-top: 0.5rem; display: ${this._state.errors.message ? 'block' : 'none'};">${this._state.errors.message || ''}</p>
                </div>

                

<div style="margin-bottom: 1.5rem;">
  <label style="display: block; margin-bottom: 0.5rem; font-weight: 500; color: #293940;">
    ${t.form.file}
  </label>
  <input 
    type="file" 
    name="file" 
    id="fileInput"
    style="display: none;" 
  />
  <div style="display: flex; gap: 10px; align-items: center; margin-top: 0.5rem;">
    <button 
      type="button" 
      id="filePickerBtn"
      style="padding: 12px 16px; background-color: #0078D4; color: white; border: none; border-radius: 8px; cursor: pointer; font-size: 1rem; transition: background-color 0.2s ease;"
      onmouseover="this.style.backgroundColor='#005A9E'"
      onmouseout="this.style.backgroundColor='#0078D4'"
    >
      ${this._state.formData.file ? 'Change File' : 'Choose File'}
    </button>
    ${this._state.formData.file ? `
      <span id="selectedFileName" style="font-size: 0.95rem; color: #293940; flex: 1; word-break: break-word;">
        ${this._state.formData.file.name}
      </span>
      <button 
        type="button" 
        id="fileRemoveBtn"
        style="padding: 6px 10px; background-color: transparent; color: #ef4444; border: 1px solid #ef4444; border-radius: 4px; cursor: pointer; font-size: 0.9rem; font-weight: bold; transition: all 0.2s ease;"
        onmouseover="this.style.backgroundColor='#ef4444'; this.style.color='white'"
        onmouseout="this.style.backgroundColor='transparent'; this.style.color='#ef4444'"
      >
        ✕
      </button>
    ` : `
      <span id="selectedFileName" style="font-size: 0.95rem; color: #8C8979;"></span>
    `}
  </div>
  <p style="color: #8C8979; font-size: 0.875rem; margin-top: 0.5rem;">
    ${t.form.fileHelp}
  </p>
  <p id="fileError" style="color: #ef4444; font-size: 0.875rem; margin-top: 0.5rem;">${this._state.errors.file ? this._state.errors.file : ''}</p>
</div>


               
<button
  type="submit"
  ${this._state.submitting ? 'disabled aria-busy="true"' : ''}
  class="submit-btn"
>
  ${this._state.submitting
        ? `
      <span class="btn-content">
        <span class="spinner" aria-hidden="true"></span>
        <span class="btn-text">${t.form.submitting}</span>
      </span>
    `
        : `
      <span class="btn-content">
        <span class="btn-text">${t.form.submitBtn}</span>
      </span>  
    `
      }
</button>

              </form>
            </div>

            <!-- FAQ -->
            <div id="faq" style="margin-bottom: 60px;">
              <h2 style="font-size: 2.5rem; font-weight: bold; text-align: center; color: #F2B29B; margin-bottom: 2rem; text-shadow: 2px 2px 4px rgba(0,0,0,0.3);">${t.faq.title}</h2>
              ${t.faq.questions.map(faq => `
                <details style="background: #F2F2F2; border-radius: 12px; padding: 20px; margin-bottom: 1rem; border: 1px solid rgba(242, 178, 155, 0.3); transition: all 0.3s ease;">
                  <summary style="font-weight: 600; font-size: 1.125rem; cursor: pointer; color: #293940;">${faq.q}</summary>
                  <p style="margin-top: 1rem; color: #3E5159; line-height: 1.6;">${faq.a}</p>
                </details>
              `).join('')}
            </div>

            <!-- Footer -->
            <div style="text-align: center; color: #F2F2F2; padding-top: 40px; border-top: 1px solid rgba(242, 178, 155, 0.3);">
              <p style="opacity: 0.9;">${t.footer}</p>
            </div>

          </div>
        </div>
      </div>
    `;
  }


  private _attachEventListeners(): void {
    if (!this._container) return;

    // --- Remove old listeners by replacing the node with a clone ---
    const oldContainer = this._container;
    const newContainer = oldContainer.cloneNode(true) as HTMLElement;
    oldContainer.parentNode?.replaceChild(newContainer, oldContainer);
    this._container = newContainer;

    // --- Handle navbar & general clicks (language toggle, menu toggle, smooth scroll) ---
    this._container.addEventListener('click', (e: Event) => {
      const target = e.target as HTMLElement;
      const action = target.getAttribute('data-action');

      if (!action) return;

      e.preventDefault();

      switch (action) {
        case 'toggleLanguage': {
          this._state.language = this._state.language === 'en' ? 'ar' : 'en';
          this._render();
          this._attachEventListeners();
          break;
        }
        case 'toggleMenu': {
          this._state.mobileMenuOpen = !this._state.mobileMenuOpen;
          this._render();
          this._attachEventListeners();
          break;
        }
        case 'scrollTo': {
          const targetId = target.getAttribute('data-target');
          if (targetId) {
            // Close menu if open, then render, then scroll smoothly
            this._state.mobileMenuOpen = false;
            this._render();
            this._attachEventListeners();
            setTimeout(() => {
              const element = document.getElementById(targetId);
              if (element) {
                element.scrollIntoView({ behavior: 'smooth', block: 'start' });
              }
            }, 100);
          }
          break;
        }
        default:
          break;
      }
    });

    // --- Form wiring ---
    const form = this._container.querySelector('#reportForm') as HTMLFormElement | null;
    if (!form) return;

    // Submit handler
    form.addEventListener('submit', (e: Event) => this._handleSubmit(e));

    // File picker button handler
    const filePickerBtn = this._container.querySelector('#filePickerBtn') as HTMLButtonElement | null;
    if (filePickerBtn) {
      filePickerBtn.addEventListener('click', (e: Event) => {
        e.preventDefault();
        const fileInput = this._container?.querySelector('#fileInput') as HTMLInputElement | null;
        if (fileInput) {
          fileInput.click();
        }
      });
    }

    // File remove button handler
    const fileRemoveBtn = this._container.querySelector('#fileRemoveBtn') as HTMLButtonElement | null;
    if (fileRemoveBtn) {
      fileRemoveBtn.addEventListener('click', (e: Event) => {
        e.preventDefault();
        this._state.formData.file = null;
        delete this._state.errors.file;
        this._render();
        this._attachEventListeners();
      });
    }

    // Input, select, and textarea handlers (including file input)
    const inputs = form.querySelectorAll('input, textarea, select');
    inputs.forEach((input) => {
      // Handle input events
      input.addEventListener('input', (e: Event) => {
        const target = e.target as HTMLInputElement;
        const fieldName = target.name as keyof typeof this._state.formData;

        // File input branch
        if (target.type === 'file') {
          const selectedFile = (target.files && target.files[0]) ? target.files[0] : null;
          this._state.formData.file = selectedFile;
          this._state.touched.add('file');

          // Always re-render on file selection to update filename, button text, and remove button
          this._render();
          this._attachEventListeners();
          return;
        }

        // Text/select/textarea branch
        // Guard against accidentally assigning a string into the `file` property
        if (fieldName !== 'file') {
          this._state.formData[fieldName] = target.value;
        }

        // Live validate for already-touched fields, but don't re-render on keystroke
        if (this._state.touched.has(String(fieldName))) {
          const error = this._validateField(String(fieldName), target.value);
          if (error) {
            this._state.errors[String(fieldName)] = error;
          } else {
            delete this._state.errors[String(fieldName)];
          }
          // Only update error message in DOM, don't re-render entire form
          const errorElement = form.querySelector(`[data-error="${fieldName}"]`) as HTMLElement | null;
          if (errorElement) {
            errorElement.textContent = error || '';
            errorElement.style.display = error ? 'block' : 'none';
          }
        }
      });

      // Handle blur (mark as touched, then validate)
      input.addEventListener('blur', (e: Event) => {
        const target = e.target as HTMLInputElement;
        const fieldName = target.name;

        this._state.touched.add(fieldName);

        // On blur, validate the field
        if (target.type === 'file') {
          const selectedFile = (target.files && target.files[0]) ? target.files[0] : this._state.formData.file;
          const error = this._validateField('file', selectedFile);
          if (error) {
            this._state.errors.file = error;
          } else {
            delete this._state.errors.file;
          }
        } else {
          const value = (this._state.formData as any)[fieldName];
          const error = this._validateField(fieldName, value);
          if (error) {
            this._state.errors[fieldName] = error;
          } else {
            delete this._state.errors[fieldName];
          }
        }

        this._render();
        this._attachEventListeners();
      });
    });
  }


  private async _handleSubmit(e: Event): Promise<void> {
    e.preventDefault();

    const t = this._getTranslations()[this._state.language];

    // Include 'file' in touched + validation pass
    const fieldsToValidate = ['email', 'message', 'site', 'management', 'file'];
    fieldsToValidate.forEach(field => {
      this._state.touched.add(field);
      const value =
        field === 'file'
          ? this._state.formData.file
          : this._state.formData[field as keyof typeof this._state.formData];

      const error = this._validateField(field, value as any);
      if (error) {
        this._state.errors[field] = error;
      } else {
        delete this._state.errors[field];
      }
    });

    if (Object.keys(this._state.errors).length > 0) {
      this._render();
      this._attachEventListeners();
      return;
    }

    // Check file size if file is present
    if (this._state.formData.file && this._state.formData.file.size > this.MAX_FILE_SIZE) {
      const maxSizeMB = this.MAX_FILE_SIZE / (1024 * 1024);
      const fileSizeMB = (this._state.formData.file.size / (1024 * 1024)).toFixed(2);
      const errorMsg = this._state.language === 'en'
        ? `File is too large (${fileSizeMB} MB). Maximum allowed size is ${maxSizeMB} MB.`
        : `الملف كبير جدًا (${fileSizeMB} MB). الحد الأقصى المسموح به هو ${maxSizeMB} MB.`;

      // Set error in state and update UI directly
      this._state.errors.file = errorMsg;
      const errEl = this._container?.querySelector('#fileError');
      if (errEl) {
        errEl.textContent = errorMsg;
      }

      this._showToast(errorMsg, 'error');
      return;
    }

    this._state.submitting = true;
    this._render();
    this._attachEventListeners();

    try {
      // Prepare optional file payload
      let filePayload: {
        name: string;
        base64: string;
      } | null = null;

      if (this._state.formData.file) {
        const f = this._state.formData.file;
        const base64 = await this._fileToBase64(f);
        filePayload = {
          name: f.name,
          base64
        };
      }

      // Normalize optional text fields to "Anonymous" if left blank
      const payload = {
        ...this._state.formData,
        name: (this._state.formData.name || '').trim() || 'Anonymous',
        email: (this._state.formData.email || '').trim() || 'Anonymous',
        department: (this._state.formData.department || '').trim() || 'Anonymous',
        // Attach file payload (or null)
        file: filePayload,
        language: this._state.language,
        timestamp: new Date().toISOString()
      };

      // Prefer property pane endpoint; fall back to hardcoded URL if not set.
      const defaultEndpoint =
        'https://default10bfe07b12fc4c4ca9122a8b161354.f9.environment.api.powerplatform.com/powerautomate/automations/direct/workflows/38d75b43da9c41aba1cb3d65326b73f9/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jziffN3HZd2PlufYFPert4xouJalQQQWyAwLFvqPshk';

      const apiEndpoint = (this.properties?.apiEndpoint || '').trim() || defaultEndpoint;

      await fetch(apiEndpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      console.log('Report payload:', payload);
      this._showToast(t.success, 'success');

      // Reset form and validation state
      this._state.formData = {
        name: '',
        email: '',
        site: '',
        department: '',
        message: '',
        file: null
      };
      this._state.touched.clear();
      this._state.errors = {};

    } catch (error) {
      console.error('Error:', error);
      this._showToast(t.errorSubmit, 'error');
    } finally {
      this._state.submitting = false;
      this._render();
      this._attachEventListeners();
    }
  }


  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }



  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }



  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('apiEndpoint', {
                  label: 'API Endpoint URL',
                  description: 'Enter the API endpoint for report submission'
                })
              ]
            }
          ]
        }
      ]
    };
  }


  private _fileToBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result as string;
        // result looks like "data:<mime>;base64,<base64>", strip the prefix:
        const base64 = result.split(',')[1] || '';
        resolve(base64);
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }


}