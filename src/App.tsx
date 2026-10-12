/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useEffect, useRef, useState } from 'react';
import {
  getIoDetails,
  getMasterTemplate,
  saveIoDetails,
  saveMasterTemplate,
  StoredTemplate,
} from './utils/db';
import { generateNoticeDocx, PlaceholderValues } from './utils/docxGenerator';

const BIRBHUM_POLICE_STATIONS = [
  'Suri PS',
  'Sainthia PS',
  'Bolpur PS',
  'Santiniketan PS',
  'Rampurhat PS',
  'Dubrajpur PS',
  'Labpur PS',
  'Nanoor PS',
  'Mayureswar PS',
  'Md. Bazar PS',
  'Murarai PS',
  'Nalhati PS',
  'Margram PS',
  'Mallarpur PS',
  'Tarapith PS',
  'Khayrasole PS',
  'Kankartala PS',
  'Rajnagar PS',
  'Illambazar PS',
  'Panrui PS',
  'Sadaipur PS',
  'Paikar PS',
  'Kirnahar PS',
  'Suri Women PS',
  'Bolpur Women PS',
  'Rampurhat Women PS',
];

function formatNoticeDate(isoDate: string): string {
  if (!isoDate) return '';
  const parts = isoDate.split('-');
  if (parts.length !== 3) return isoDate;
  const [year, month, day] = parts;
  return `${day}.${month}.${year}`;
}

export default function App() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const ioPopoverRef = useRef<HTMLDivElement | null>(null);
  const masterPopoverRef = useRef<HTMLDivElement | null>(null);

  const [masterTemplate, setMasterTemplate] = useState<StoredTemplate | null>(null);
  const [statusMessage, setStatusMessage] = useState<{
    type: 'error' | 'info';
    text: string;
  } | null>(null);

  // IO Details (Saved permanently in IndexedDB / localStorage)
  const [ioName, setIoName] = useState('');
  const [ioMobile, setIoMobile] = useState('');
  const [isIoOpen, setIsIoOpen] = useState(false);
  const [draftIoName, setDraftIoName] = useState('');
  const [draftIoMobile, setDraftIoMobile] = useState('');

  // Master Template popover state
  const [isMasterOpen, setIsMasterOpen] = useState(false);

  // PWA Install Prompt state
  interface BeforeInstallPromptEvent extends Event {
    prompt: () => Promise<void>;
    userChoice: Promise<{ outcome: 'accepted' | 'dismissed'; platform: string }>;
  }
  const [deferredInstallPrompt, setDeferredInstallPrompt] =
    useState<BeforeInstallPromptEvent | null>(null);
  const [isAppInstalled, setIsAppInstalled] = useState(false);

  // Notice Form Fields
  const [memoNo, setMemoNo] = useState('');
  const [noticeDateIso, setNoticeDateIso] = useState('');
  const [gdeNo, setGdeNo] = useState('');
  const [gdeDateIso, setGdeDateIso] = useState('');
  const [ncmecTiplineNo, setNcmecTiplineNo] = useState('');
  const [addresseeName, setAddresseeName] = useState('');
  const [fullAddress, setFullAddress] = useState('');
  const [addresseePs, setAddresseePs] = useState('');
  const [authority, setAuthority] = useState<'O/C' | 'I/C'>('I/C');

  const trimmedPs = addresseePs.trim();

  useEffect(() => {
    getMasterTemplate()
      .then((stored) => {
        if (stored) {
          setMasterTemplate(stored);
        }
      })
      .catch(() => {
        setStatusMessage({
          type: 'error',
          text: 'Could not read stored master template from local database.',
        });
      });

    getIoDetails().then((savedIo) => {
      setIoName(savedIo.ioName);
      setIoMobile(savedIo.ioMobile);
      setDraftIoName(savedIo.ioName);
      setDraftIoMobile(savedIo.ioMobile);
    });

    const isStandalone =
      window.matchMedia('(display-mode: standalone)').matches ||
      (window.navigator as unknown as { standalone?: boolean }).standalone === true;
    if (isStandalone) {
      setIsAppInstalled(true);
    }

    const handleBeforeInstallPrompt = (e: Event) => {
      e.preventDefault();
      setDeferredInstallPrompt(e as BeforeInstallPromptEvent);
    };

    const handleAppInstalled = () => {
      setIsAppInstalled(true);
      setDeferredInstallPrompt(null);
    };

    window.addEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
    window.addEventListener('appinstalled', handleAppInstalled);

    return () => {
      window.removeEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
      window.removeEventListener('appinstalled', handleAppInstalled);
    };
  }, []);

  const handleInstallApp = async () => {
    if (!deferredInstallPrompt) {
      return;
    }
    try {
      await deferredInstallPrompt.prompt();
      const { outcome } = await deferredInstallPrompt.userChoice;
      if (outcome === 'accepted') {
        setIsAppInstalled(true);
        setDeferredInstallPrompt(null);
      }
    } catch {
      // If prompt cannot be shown, keep the app working normally without showing an error
    }
  };

  useEffect(() => {
    if (!isIoOpen && !isMasterOpen) return;
    const handleClickOutside = (event: MouseEvent) => {
      if (
        isIoOpen &&
        ioPopoverRef.current &&
        !ioPopoverRef.current.contains(event.target as Node)
      ) {
        setIsIoOpen(false);
      }
      if (
        isMasterOpen &&
        masterPopoverRef.current &&
        !masterPopoverRef.current.contains(event.target as Node)
      ) {
        setIsMasterOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [isIoOpen, isMasterOpen]);

  const handleToggleIoPopover = () => {
    if (!isIoOpen) {
      setDraftIoName(ioName);
      setDraftIoMobile(ioMobile);
      setIsMasterOpen(false);
    }
    setIsIoOpen((prev) => !prev);
  };

  const handleToggleMasterPopover = () => {
    if (!isMasterOpen) {
      setIsIoOpen(false);
    }
    setIsMasterOpen((prev) => !prev);
  };

  const handleSaveIoDetails = async (e: React.FormEvent) => {
    e.preventDefault();
    const trimmedName = draftIoName.trim();
    const trimmedMobile = draftIoMobile.trim();

    if (!trimmedName || !trimmedMobile) {
      setStatusMessage({
        type: 'error',
        text: 'Please enter both IO Name and IO Mobile Number before saving.',
      });
      return;
    }

    try {
      await saveIoDetails(trimmedName, trimmedMobile);
      setIoName(trimmedName);
      setIoMobile(trimmedMobile);
      setIsIoOpen(false);
      setStatusMessage({
        type: 'info',
        text: 'IO Details saved.',
      });
    } catch {
      setStatusMessage({
        type: 'error',
        text: 'Failed to save IO Details.',
      });
    }
  };

  const handleMasterUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!file.name.toLowerCase().endsWith('.docx')) {
      setStatusMessage({
        type: 'error',
        text: 'Please upload a valid .docx file for the 179 BNSS NCMEC Notice master template.',
      });
      e.target.value = '';
      return;
    }

    try {
      const saved = await saveMasterTemplate(file);
      setMasterTemplate(saved);
      setIsMasterOpen(false);
      setStatusMessage({
        type: 'info',
        text: `179 BNSS NCMEC Notice master ("${saved.fileName}") saved locally.`,
      });
    } catch {
      setStatusMessage({
        type: 'error',
        text: 'Failed to save master template to IndexedDB.',
      });
    } finally {
      e.target.value = '';
    }
  };

  const handleGenerateNotice = (e: React.FormEvent) => {
    e.preventDefault();
    setStatusMessage(null);

    if (!masterTemplate) {
      setIsMasterOpen(true);
      setStatusMessage({
        type: 'error',
        text: 'Master template missing. Please upload the "179 BNSS NCMEC Notice" (.docx) master template first.',
      });
      return;
    }

    if (!ioName.trim() || !ioMobile.trim()) {
      setDraftIoName(ioName);
      setDraftIoMobile(ioMobile);
      setIsIoOpen(true);
      setStatusMessage({
        type: 'error',
        text: 'IO Details missing. Please enter and save IO Name and IO Mobile Number in "IO Details" at the top-right.',
      });
      return;
    }

    const formattedNoticeDate = formatNoticeDate(noticeDateIso);
    const formattedGdeDate = formatNoticeDate(gdeDateIso);

    const missingFields: string[] = [];
    if (!memoNo.trim()) missingFields.push('Memo No.');
    if (!formattedNoticeDate) missingFields.push('Notice Date');
    if (!gdeNo.trim()) missingFields.push('GDE No.');
    if (!formattedGdeDate) missingFields.push('GDE Date');
    if (!ncmecTiplineNo.trim()) missingFields.push('NCMEC Tipline Report No.');
    if (!addresseeName.trim()) missingFields.push('Addressee Name');
    if (!fullAddress.trim()) missingFields.push('Full Address');
    if (!trimmedPs) missingFields.push('Addressee PS');
    if (!authority) missingFields.push('Authority');

    if (missingFields.length > 0) {
      setStatusMessage({
        type: 'error',
        text: `Please fill all required fields: ${missingFields.join(', ')}.`,
      });
      return;
    }

    try {
      // Do not modify or rewrite the pasted address
      const placeholderValues: PlaceholderValues = {
        MEMO_NO: memoNo.trim(),
        DATE: formattedNoticeDate,
        GDE_NO: gdeNo.trim(),
        GDE_DATE: formattedGdeDate,
        NCMEC_TIPLINE_NO: ncmecTiplineNo.trim(),
        PS_AUTHORITY: authority,
        ADDRESSEE_PS: trimmedPs,
        ADDRESSEE_NAME: addresseeName.trim(),
        FULL_ADDRESS: fullAddress,
        IO_NAME: ioName.trim(),
        IO_MOBILE: ioMobile.trim(),
      };

      const blob = generateNoticeDocx(masterTemplate.buffer, placeholderValues);

      const safeName = (addresseeName.trim() || 'Notice').replace(/[^a-zA-Z0-9_-]+/g, '_');
      const downloadName = `179_BNSS_NCMEC_Notice_${safeName}.docx`;

      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = downloadName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (err) {
      const message =
        err instanceof Error
          ? err.message
          : 'Unable to generate notice from the uploaded DOCX master.';
      setStatusMessage({
        type: 'error',
        text: message,
      });
    }
  };

  const inputClassName =
    'w-full h-9 px-3 text-sm text-[#0F172A] bg-white border border-slate-300 rounded focus:outline-none focus:border-[#0F172A] focus:ring-1 focus:ring-[#0F172A]';
  const labelClassName = 'block text-xs font-medium text-slate-700 mb-1.5';

  return (
    <div className="min-h-screen bg-[#F8FAFC] text-[#0F172A] flex flex-col">
      {/* Top Bar */}
      <header className="w-full bg-white border-b border-slate-200">
        <div className="max-w-[850px] mx-auto px-6 py-4 flex items-center justify-between gap-4">
          <div>
            <h1 className="text-base font-semibold tracking-tight text-[#0F172A] leading-snug">
              179 Notice Generator
            </h1>
            <p className="text-xs text-slate-600 leading-tight mt-0.5">
              Cyber Crime PS, Birbhum
            </p>
          </div>

          <div className="flex items-center gap-2.5 shrink-0">
            {/* Install App */}
            {!isAppInstalled && (
              <button
                type="button"
                onClick={handleInstallApp}
                className="px-3.5 py-1.5 text-xs font-medium text-[#0F172A] bg-white border border-slate-300 rounded hover:bg-slate-50 focus:outline-none focus:ring-1 focus:ring-[#0F172A] cursor-pointer whitespace-nowrap"
              >
                Install App
              </button>
            )}

            {/* IO Details */}
            <div className="relative" ref={ioPopoverRef}>
              <button
                type="button"
                onClick={handleToggleIoPopover}
                className="px-3.5 py-1.5 text-xs font-medium text-[#0F172A] bg-white border border-slate-300 rounded hover:bg-slate-50 focus:outline-none focus:ring-1 focus:ring-[#0F172A] cursor-pointer whitespace-nowrap"
              >
                IO Details
              </button>

              {isIoOpen && (
                <div className="absolute right-0 mt-2 w-72 bg-white border border-slate-200 rounded-md p-4 z-20">
                  <form onSubmit={handleSaveIoDetails} className="space-y-3">
                    <div>
                      <label htmlFor="ioNameSetting" className={labelClassName}>
                        IO Name
                      </label>
                      <input
                        id="ioNameSetting"
                        type="text"
                        value={draftIoName}
                        onChange={(e) => setDraftIoName(e.target.value)}
                        className={inputClassName}
                      />
                    </div>

                    <div>
                      <label htmlFor="ioMobileSetting" className={labelClassName}>
                        IO Mobile Number
                      </label>
                      <input
                        id="ioMobileSetting"
                        type="tel"
                        value={draftIoMobile}
                        onChange={(e) => setDraftIoMobile(e.target.value)}
                        className={`${inputClassName} tabular-nums`}
                      />
                    </div>

                    <div className="pt-1">
                      <button
                        type="submit"
                        className="w-full px-3.5 py-1.5 text-xs font-medium text-white bg-[#0F172A] rounded hover:bg-slate-800 focus:outline-none focus:ring-1 focus:ring-[#0F172A] cursor-pointer"
                      >
                        Save
                      </button>
                    </div>
                  </form>
                </div>
              )}
            </div>

            {/* Master Template */}
            <div className="relative" ref={masterPopoverRef}>
              <input
                ref={fileInputRef}
                type="file"
                accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                onChange={handleMasterUpload}
                className="hidden"
              />
              <button
                type="button"
                onClick={handleToggleMasterPopover}
                className="px-3.5 py-1.5 text-xs font-medium text-[#0F172A] bg-white border border-slate-300 rounded hover:bg-slate-50 focus:outline-none focus:ring-1 focus:ring-[#0F172A] cursor-pointer whitespace-nowrap"
              >
                Master Template
              </button>

              {isMasterOpen && (
                <div className="absolute right-0 mt-2 w-72 bg-white border border-slate-200 rounded-md p-4 z-20 space-y-3">
                  <div>
                    <div className="text-xs font-semibold text-[#0F172A]">
                      179 BNSS NCMEC Notice
                    </div>
                    <div className="text-xs text-slate-600 mt-1 truncate">
                      {masterTemplate ? masterTemplate.fileName : 'No master uploaded'}
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={() => fileInputRef.current?.click()}
                    className="w-full px-3.5 py-1.5 text-xs font-medium text-white bg-[#0F172A] rounded hover:bg-slate-800 focus:outline-none focus:ring-1 focus:ring-[#0F172A] cursor-pointer"
                  >
                    {masterTemplate ? 'Replace Master (.docx)' : 'Upload Master (.docx)'}
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 w-full px-4 sm:px-6 py-8 sm:py-10">
        <div className="max-w-[850px] mx-auto bg-white border border-slate-200 rounded-md p-6 sm:p-8">
          {statusMessage && (
            <div
              role="status"
              className={`mb-6 px-3.5 py-2.5 text-xs rounded border ${
                statusMessage.type === 'error'
                  ? 'bg-red-50 border-red-200 text-red-900'
                  : 'bg-slate-50 border-slate-200 text-[#0F172A]'
              }`}
            >
              {statusMessage.text}
            </div>
          )}

          <form onSubmit={handleGenerateNotice} className="space-y-5">
            {/* Row 1: Memo No. & Notice Date */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div>
                <label htmlFor="memoNo" className={labelClassName}>
                  Memo No.
                </label>
                <input
                  id="memoNo"
                  type="text"
                  value={memoNo}
                  onChange={(e) => setMemoNo(e.target.value)}
                  className={`${inputClassName} tabular-nums`}
                />
              </div>

              <div>
                <label htmlFor="noticeDate" className={labelClassName}>
                  Notice Date
                </label>
                <input
                  id="noticeDate"
                  type="date"
                  value={noticeDateIso}
                  onChange={(e) => setNoticeDateIso(e.target.value)}
                  className={`${inputClassName} tabular-nums`}
                />
              </div>
            </div>

            {/* Row 2: GDE No., GDE Date, NCMEC Tipline Report No. */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div>
                <label htmlFor="gdeNo" className={labelClassName}>
                  GDE No.
                </label>
                <input
                  id="gdeNo"
                  type="text"
                  value={gdeNo}
                  onChange={(e) => setGdeNo(e.target.value)}
                  className={`${inputClassName} tabular-nums`}
                />
              </div>

              <div>
                <label htmlFor="gdeDate" className={labelClassName}>
                  GDE Date
                </label>
                <input
                  id="gdeDate"
                  type="date"
                  value={gdeDateIso}
                  onChange={(e) => setGdeDateIso(e.target.value)}
                  className={`${inputClassName} tabular-nums`}
                />
              </div>

              <div>
                <label htmlFor="ncmecTiplineNo" className={labelClassName}>
                  NCMEC Tipline Report No.
                </label>
                <input
                  id="ncmecTiplineNo"
                  type="text"
                  value={ncmecTiplineNo}
                  onChange={(e) => setNcmecTiplineNo(e.target.value)}
                  className={`${inputClassName} tabular-nums`}
                />
              </div>
            </div>

            {/* Row 3: Addressee Name */}
            <div>
              <label htmlFor="addresseeName" className={labelClassName}>
                Addressee Name
              </label>
              <input
                id="addresseeName"
                type="text"
                value={addresseeName}
                onChange={(e) => setAddresseeName(e.target.value)}
                className={inputClassName}
              />
            </div>

            {/* Row 4: Full Address */}
            <div>
              <label htmlFor="fullAddress" className={labelClassName}>
                Full Address
              </label>
              <textarea
                id="fullAddress"
                rows={4}
                value={fullAddress}
                onChange={(e) => setFullAddress(e.target.value)}
                className="w-full p-3 text-sm text-[#0F172A] bg-white border border-slate-300 rounded focus:outline-none focus:border-[#0F172A] focus:ring-1 focus:ring-[#0F172A]"
              />
            </div>

            {/* Row 5: Addressee PS & Authority */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div className="sm:col-span-2">
                <label htmlFor="addresseePs" className={labelClassName}>
                  Addressee PS
                </label>
                <input
                  id="addresseePs"
                  type="text"
                  list="birbhum-ps-options"
                  value={addresseePs}
                  onChange={(e) => setAddresseePs(e.target.value)}
                  className={inputClassName}
                />
                <datalist id="birbhum-ps-options">
                  {BIRBHUM_POLICE_STATIONS.map((ps) => (
                    <option key={ps} value={ps} />
                  ))}
                </datalist>
              </div>

              <div>
                <label htmlFor="authority" className={labelClassName}>
                  Authority
                </label>
                <select
                  id="authority"
                  value={authority}
                  onChange={(e) => setAuthority(e.target.value as 'O/C' | 'I/C')}
                  className={inputClassName}
                >
                  <option value="O/C">O/C</option>
                  <option value="I/C">I/C</option>
                </select>
              </div>
            </div>

            {/* Primary Button */}
            <div className="pt-2">
              <button
                type="submit"
                className="w-full sm:w-auto px-6 py-2.5 text-sm font-medium text-white bg-[#0F172A] rounded hover:bg-slate-800 focus:outline-none focus:ring-2 focus:ring-[#0F172A] focus:ring-offset-2 cursor-pointer"
              >
                Generate Notice
              </button>
            </div>
          </form>
        </div>
      </main>
    </div>
  );
}


