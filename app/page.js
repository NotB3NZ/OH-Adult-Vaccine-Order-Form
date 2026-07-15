"use client";

import { useState, useCallback, useMemo } from "react";
import {
  ShieldCheck,
  Plus,
  Minus,
  ShoppingCart,
  UserPlus,
  MapPin,
  Send,
  ChevronDown,
  Syringe,
  AlertCircle,
  CheckCircle2,
  Loader2,
  X,
  Info,
  Package,
  User,
  Users,
} from "lucide-react";

/* ================================================================
   CONSTANTS
   ================================================================ */

const VACCINES = [
  { id: "DPT", name: "DPT", brand: "Boostrix", price: 1210, maxQty: 10, colKey: "DPT" },
  { id: "HepA", name: "Hepatitis A", brand: "Mevac A", price: 1410, maxQty: 10, colKey: "HepA" },
  { id: "HepB", name: "Hepatitis B", brand: "Revac B", price: 660, maxQty: 10, colKey: "HepB" },
  { id: "HepAB", name: "Hepatitis A and B", brand: "Twinrix Adult", price: 1980, maxQty: 10, colKey: "HepAB" },
  { id: "HPV", name: "HPV", brand: "Gardasil Nona", price: 5300, maxQty: 10, colKey: "HPV" },
  { id: "MMR", name: "MMR", brand: "Tresivac", price: 638, maxQty: 10, colKey: "MMR" },
  { id: "Pneumo13", name: "Pneumococcal", brand: "Prevenar 13", price: 2700, maxQty: 10, colKey: "Pneumo13" },
  { id: "Pneumo20", name: "Pneumococcal", brand: "Prevenar 20", price: 3680, maxQty: 10, colKey: "Pneumo20" },
  { id: "Pneumo23", name: "Pneumococcal", brand: "Pneumovax23", price: 1870, maxQty: 10, colKey: "Pneumo23" },
  { id: "Shingles", name: "Shingles", brand: "Shingrix", price: 6700, maxQty: 10, colKey: "Shingles" },
  { id: "Varicella", name: "Varicella", brand: "Varivax", price: 1055, maxQty: 10, colKey: "Varicella" },
  { id: "RabiesSpeeda", name: "Rabies", brand: "Speeda", price: 950, maxQty: 1, colKey: "RabiesSpeeda" },
  { id: "RabiesVerorab", name: "Rabies", brand: "VERORAB", price: 1920, maxQty: 1, colKey: "RabiesVerorab" },
];

// TODO: Replace with your actual delivery site options
const DELIVERY_SITES = [
  "BGC Support Center",
  "Davao Plant 1",
  "Davao Plant 2",
  "Misamis Occidental Plant",
  "Calasiao Plant",
];

const CONSENT_TEXT = `loren ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.`;

/* ================================================================
   HELPER: Generate a simple unique ID
   ================================================================ */
function generateOrderId() {
  const timestamp = Date.now().toString(36).toUpperCase();
  const random = Math.random().toString(36).substring(2, 8).toUpperCase();
  return `ORD-${timestamp}-${random}`;
}

/* ================================================================
   HELPER: Format PHP currency
   ================================================================ */
function formatPrice(amount) {
  return `₱${amount.toLocaleString("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

/* ================================================================
   HELPER: Create empty basket
   ================================================================ */
function createEmptyBasket() {
  const basket = {};
  VACCINES.forEach((v) => {
    basket[v.id] = { qty: 0, siteOverride: "" };
  });
  return basket;
}

/* ================================================================
   HELPER: Calculate basket total
   ================================================================ */
function calculateTotal(basket) {
  return VACCINES.reduce((sum, v) => sum + basket[v.id].qty * v.price, 0);
}

/* ================================================================
   HELPER: Count basket items
   ================================================================ */
function countItems(basket) {
  return VACCINES.reduce((sum, v) => sum + basket[v.id].qty, 0);
}

/* ================================================================
   CORE: prepareExcelPayload
   Flattens the order state into an array of flat row objects
   compatible with Excel Cloud / Microsoft Graph API.
   ================================================================ */
function prepareExcelPayload({
  associateNumber,
  defaultSite,
  consentAgreed,
  associateBasket,
  dependentBasket,
  wantDependent,
}) {
  const orderId = generateOrderId();
  const timestamp = new Date().toISOString();
  const rows = [];

  function buildRow(basket, orderType) {
    const row = {
      OrderID: orderId,
      Timestamp: timestamp,
      AssociateNumber: associateNumber,
      OrderType: orderType,
    };

    VACCINES.forEach((v) => {
      const item = basket[v.id];
      row[`${v.colKey}_Qty`] = item.qty;
      row[`${v.colKey}_Site`] = item.qty > 0
        ? (item.siteOverride || defaultSite)
        : "";
    });

    row.TotalCost = calculateTotal(basket);
    row.ConsentAgreed = consentAgreed ? "Yes" : "No";

    return row;
  }

  // Always include associate order if there are items
  if (countItems(associateBasket) > 0) {
    rows.push(buildRow(associateBasket, "Associate"));
  }

  // Include dependent order if toggled on and has items
  if (wantDependent && countItems(dependentBasket) > 0) {
    rows.push(buildRow(dependentBasket, "Dependent"));
  }

  return rows;
}


/* ================================================================
   COMPONENT: VaccineCard
   ================================================================ */
function VaccineCard({ vaccine, item, defaultSite, onQtyChange, onSiteChange }) {
  const isActive = item.qty > 0;
  const isMaxed = item.qty >= vaccine.maxQty;
  const isRabies = vaccine.maxQty === 1;

  return (
    <div
      className={`vaccine-card rounded-2xl border-2 p-5 bg-card-bg ${isActive ? "active border-primary/40" : "border-card-border"
        }`}
      style={{ boxShadow: "var(--shadow-sm)" }}
    >
      {/* Header */}
      <div className="flex items-start justify-between gap-3 mb-3">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2 mb-1">
            <Syringe size={18} className="text-primary shrink-0" />
            <h3 className="text-lg font-semibold text-foreground truncate">
              {vaccine.name}
            </h3>
          </div>
          <p className="text-sm text-muted">{vaccine.brand}</p>
        </div>
        <div className="text-right shrink-0">
          <p className="text-lg font-bold text-primary">
            {formatPrice(vaccine.price)}
          </p>
          <p className="text-xs text-muted">per dose</p>
        </div>
      </div>

      {/* Max dose notice for Rabies */}
      {isRabies && (
        <div className="flex items-center gap-1.5 mb-3 text-xs text-warning bg-warning-light rounded-lg px-3 py-1.5">
          <AlertCircle size={14} />
          <span>Maximum 1 dose</span>
        </div>
      )}

      {/* Quantity Selector */}
      <div className="flex items-center justify-between bg-muted-bg rounded-xl px-3 py-2">
        <span className="text-sm font-medium text-muted">Quantity</span>
        <div className="flex items-center gap-3">
          <button
            type="button"
            onClick={() => onQtyChange(Math.max(0, item.qty - 1))}
            disabled={item.qty === 0}
            className="qty-btn w-9 h-9 flex items-center justify-center rounded-lg bg-card-bg border border-card-border text-foreground disabled:opacity-30 disabled:cursor-not-allowed hover:bg-danger-light hover:border-danger hover:text-danger"
            aria-label={`Decrease ${vaccine.name} quantity`}
          >
            <Minus size={16} />
          </button>
          <span
            className={`text-xl font-bold w-8 text-center tabular-nums ${isActive ? "text-primary" : "text-muted"
              }`}
          >
            {item.qty}
          </span>
          <button
            type="button"
            onClick={() => onQtyChange(Math.min(vaccine.maxQty, item.qty + 1))}
            disabled={isMaxed}
            className="qty-btn w-9 h-9 flex items-center justify-center rounded-lg bg-card-bg border border-card-border text-foreground disabled:opacity-30 disabled:cursor-not-allowed hover:bg-primary-light hover:border-primary hover:text-primary"
            aria-label={`Increase ${vaccine.name} quantity`}
          >
            <Plus size={16} />
          </button>
        </div>
      </div>

      {/* Subtotal when active */}
      {isActive && (
        <div className="mt-3 flex items-center justify-between text-sm animate-fade-in">
          <span className="text-muted">Subtotal</span>
          <span className="font-bold text-primary">
            {formatPrice(item.qty * vaccine.price)}
          </span>
        </div>
      )}

      {/* Delivery Site Override */}
      {isActive && (
        <div className="mt-3 animate-fade-in">
          <div className="flex items-center gap-1.5 mb-1.5">
            <MapPin size={14} className="text-muted" />
            <span className="text-xs text-muted">Delivery site</span>
          </div>
          <select
            value={item.siteOverride}
            onChange={(e) => onSiteChange(e.target.value)}
            className="w-full text-sm rounded-lg border border-card-border bg-card-bg px-3 py-2 text-foreground appearance-none cursor-pointer"
            aria-label={`Delivery site for ${vaccine.name}`}
          >
            <option value="">
              Default ({defaultSite || "Select above"})
            </option>
            {DELIVERY_SITES.map((site) => (
              <option key={site} value={site}>
                {site}
              </option>
            ))}
          </select>
        </div>
      )}
    </div>
  );
}


/* ================================================================
   COMPONENT: BasketGrid
   ================================================================ */
function BasketGrid({ basket, defaultSite, onUpdate }) {
  const handleQtyChange = useCallback(
    (vaccineId, qty) => {
      onUpdate((prev) => ({
        ...prev,
        [vaccineId]: { ...prev[vaccineId], qty },
      }));
    },
    [onUpdate]
  );

  const handleSiteChange = useCallback(
    (vaccineId, site) => {
      onUpdate((prev) => ({
        ...prev,
        [vaccineId]: { ...prev[vaccineId], siteOverride: site },
      }));
    },
    [onUpdate]
  );

  return (
    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
      {VACCINES.map((vaccine) => (
        <VaccineCard
          key={vaccine.id}
          vaccine={vaccine}
          item={basket[vaccine.id]}
          defaultSite={defaultSite}
          onQtyChange={(qty) => handleQtyChange(vaccine.id, qty)}
          onSiteChange={(site) => handleSiteChange(vaccine.id, site)}
        />
      ))}
    </div>
  );
}


/* ================================================================
   COMPONENT: OrderSummary
   ================================================================ */
function OrderSummary({ associateBasket, dependentBasket, wantDependent }) {
  const associateTotal = calculateTotal(associateBasket);
  const dependentTotal = wantDependent ? calculateTotal(dependentBasket) : 0;
  const grandTotal = associateTotal + dependentTotal;
  const associateItems = countItems(associateBasket);
  const dependentItems = wantDependent ? countItems(dependentBasket) : 0;

  if (associateItems === 0 && dependentItems === 0) {
    return (
      <div className="text-center py-10 text-muted">
        <ShoppingCart size={48} className="mx-auto mb-3 opacity-40" />
        <p className="text-lg">Your baskets are empty</p>
        <p className="text-sm mt-1">Add vaccines above to see your order summary</p>
      </div>
    );
  }

  function renderBasketSummary(basket, label, icon, itemCount, total) {
    if (itemCount === 0) return null;
    const activeItems = VACCINES.filter((v) => basket[v.id].qty > 0);

    return (
      <div className="animate-fade-in">
        <div className="flex items-center gap-2 mb-3">
          {icon}
          <h4 className="font-semibold text-foreground">{label}</h4>
          <span className="ml-auto text-sm bg-primary-light text-primary-dark px-2.5 py-0.5 rounded-full font-medium">
            {itemCount} {itemCount === 1 ? "item" : "items"}
          </span>
        </div>
        <div className="space-y-2">
          {activeItems.map((v) => (
            <div
              key={v.id}
              className="flex items-center justify-between text-sm px-3 py-2 rounded-lg bg-muted-bg"
            >
              <div>
                <span className="font-medium text-foreground">{v.name}</span>
                <span className="text-muted ml-1">({v.brand})</span>
                <span className="text-muted ml-2">×{basket[v.id].qty}</span>
              </div>
              <span className="font-semibold text-foreground">
                {formatPrice(basket[v.id].qty * v.price)}
              </span>
            </div>
          ))}
        </div>
        <div className="flex justify-between items-center mt-3 pt-3 border-t border-divider">
          <span className="font-medium text-muted">Subtotal</span>
          <span className="text-lg font-bold text-foreground">
            {formatPrice(total)}
          </span>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      {renderBasketSummary(
        associateBasket,
        "My Basket (Associate)",
        <User size={18} className="text-primary" />,
        associateItems,
        associateTotal
      )}

      {wantDependent &&
        renderBasketSummary(
          dependentBasket,
          "Dependent's Basket",
          <Users size={18} className="text-accent" />,
          dependentItems,
          dependentTotal
        )}

      {/* Grand Total */}
      <div className="bg-gradient-to-r from-primary to-primary-dark rounded-2xl p-5 text-white">
        <div className="flex items-center justify-between">
          <div>
            <p className="text-sm opacity-80">Grand Total</p>
            <p className="text-3xl font-bold mt-1">
              {formatPrice(grandTotal)}
            </p>
          </div>
          <Package size={36} className="opacity-60" />
        </div>
      </div>
    </div>
  );
}


/* ================================================================
   MAIN PAGE
   ================================================================ */
export default function OrderFormPage() {
  // -- Core state --
  const [consentAgreed, setConsentAgreed] = useState(false);
  const [associateNumber, setAssociateNumber] = useState("");
  const [defaultSite, setDefaultSite] = useState("");
  const [wantDependent, setWantDependent] = useState(false);
  const [activeTab, setActiveTab] = useState("associate");

  // -- Baskets --
  const [associateBasket, setAssociateBasket] = useState(createEmptyBasket);
  const [dependentBasket, setDependentBasket] = useState(createEmptyBasket);

  // -- Submit state --
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submitResult, setSubmitResult] = useState(null); // { success, message }

  // -- Derived --
  const associateItemCount = useMemo(() => countItems(associateBasket), [associateBasket]);
  const dependentItemCount = useMemo(() => countItems(dependentBasket), [dependentBasket]);
  const totalItemCount = associateItemCount + (wantDependent ? dependentItemCount : 0);
  const grandTotal = useMemo(
    () => calculateTotal(associateBasket) + (wantDependent ? calculateTotal(dependentBasket) : 0),
    [associateBasket, dependentBasket, wantDependent]
  );

  const canSubmit = consentAgreed && associateNumber.trim() && defaultSite && totalItemCount > 0;

  // -- Handlers --
  const handleSubmit = useCallback(async () => {
    if (!canSubmit) return;

    setIsSubmitting(true);
    setSubmitResult(null);

    const payload = prepareExcelPayload({
      associateNumber,
      defaultSite,
      consentAgreed,
      associateBasket,
      dependentBasket,
      wantDependent,
    });

    try {
      const res = await fetch("/api/submit-to-excel", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ rows: payload }),
      });

      const data = await res.json();

      if (res.ok) {
        setSubmitResult({
          success: true,
          message: `Order submitted successfully! Order ID: ${payload[0]?.OrderID}`,
        });
        // Reset baskets on success
        setAssociateBasket(createEmptyBasket());
        setDependentBasket(createEmptyBasket());
        setWantDependent(false);
        setAssociateNumber("");
        setDefaultSite("");
        setConsentAgreed(false);
        setActiveTab("associate");
      } else {
        setSubmitResult({
          success: false,
          message: data.error || "Something went wrong. Please try again.",
        });
      }
    } catch {
      setSubmitResult({
        success: false,
        message: "Network error. Please check your connection and try again.",
      });
    } finally {
      setIsSubmitting(false);
    }
  }, [canSubmit, associateNumber, defaultSite, consentAgreed, associateBasket, dependentBasket, wantDependent]);

  const handleReset = useCallback(() => {
    setAssociateBasket(createEmptyBasket());
    setDependentBasket(createEmptyBasket());
  }, []);

  return (
    <main className="flex-1 pb-20">
      {/* ============================================================
          HEADER
          ============================================================ */}
      <header className="sticky top-0 z-50 bg-white/80 backdrop-blur-lg border-b border-divider">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-11 h-11 rounded-xl bg-gradient-to-br from-primary to-primary-dark flex items-center justify-center">
              <Syringe size={22} className="text-white" />
            </div>
            <div>
              <h1 className="text-xl sm:text-2xl font-bold text-foreground leading-tight">
                Adult Vaccines Order Form
              </h1>
              <p className="text-xs sm:text-sm text-muted">
                Quick &amp; easy vaccine ordering
              </p>
            </div>
          </div>
          {totalItemCount > 0 && (
            <div className="flex items-center gap-2 bg-primary-light text-primary-dark px-4 py-2 rounded-full font-medium text-sm animate-fade-in">
              <ShoppingCart size={18} />
              <span>
                {totalItemCount} {totalItemCount === 1 ? "item" : "items"}
              </span>
              <span className="hidden sm:inline">
                · {formatPrice(grandTotal)}
              </span>
            </div>
          )}
        </div>
      </header>

      <div className="max-w-6xl mx-auto px-4 sm:px-6 mt-6 space-y-6">
        {/* ============================================================
            CONSENT SECTION
            ============================================================ */}
        <section
          className="rounded-2xl border-2 border-card-border bg-card-bg p-6 sm:p-8"
          style={{ boxShadow: "var(--shadow-sm)" }}
        >
          <div className="flex items-start gap-3 mb-4">
            <ShieldCheck
              size={28}
              className={consentAgreed ? "text-success" : "text-muted"}
            />
            <div>
              <h2 className="text-xl font-bold text-foreground">
                Consent &amp; Agreement
              </h2>
              <p className="text-sm text-muted mt-1">
                Please read and accept before proceeding
              </p>
            </div>
          </div>

          <div className="bg-muted-bg rounded-xl p-5 text-sm text-foreground leading-relaxed whitespace-pre-line max-h-48 overflow-y-auto scrollbar-thin mb-5">
            {CONSENT_TEXT}
          </div>

          <label className="flex items-center gap-3 cursor-pointer group select-none">
            <input
              type="checkbox"
              checked={consentAgreed}
              onChange={(e) => setConsentAgreed(e.target.checked)}
              className="shrink-0"
              id="consent-checkbox"
            />
            <span className="text-base font-medium text-foreground group-hover:text-primary transition-colors">
              I have read and agree to the terms above
            </span>
            {consentAgreed && (
              <CheckCircle2 size={20} className="text-success ml-auto animate-checkmark" />
            )}
          </label>
        </section>

        {/* ============================================================
            GATED CONTENT (only interactive when consent is given)
            ============================================================ */}
        <div className={`consent-gate ${consentAgreed ? "unlocked" : ""}`}>
          <div className="space-y-6">
            {/* ----------------------------------------------------------
                ASSOCIATE INFO
                ---------------------------------------------------------- */}
            <section
              className="rounded-2xl border-2 border-card-border bg-card-bg p-6 sm:p-8"
              style={{ boxShadow: "var(--shadow-sm)" }}
            >
              <h2 className="text-xl font-bold text-foreground mb-5 flex items-center gap-2">
                <User size={22} className="text-primary" />
                Your Information
              </h2>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-5">
                {/* Associate Number */}
                <div>
                  <label
                    htmlFor="associate-number"
                    className="block text-sm font-medium text-foreground mb-2"
                  >
                    Associate Number <span className="text-danger">*</span>
                  </label>
                  <input
                    type="text"
                    id="associate-number"
                    value={associateNumber}
                    onChange={(e) => setAssociateNumber(e.target.value)}
                    placeholder="e.g. 12345"
                    className="w-full rounded-xl border-2 border-card-border bg-card-bg px-4 py-3 text-base text-foreground placeholder:text-muted"
                  />
                </div>

                {/* Default Delivery Site */}
                <div>
                  <label
                    htmlFor="default-site"
                    className="block text-sm font-medium text-foreground mb-2"
                  >
                    Associate Site <span className="text-danger">*</span>
                  </label>
                  <div className="relative">
                    <select
                      id="default-site"
                      value={defaultSite}
                      onChange={(e) => setDefaultSite(e.target.value)}
                      className="w-full rounded-xl border-2 border-card-border bg-card-bg px-4 py-3 text-base text-foreground appearance-none cursor-pointer"
                    >
                      <option value="">-- Select a site --</option>
                      {DELIVERY_SITES.map((site) => (
                        <option key={site} value={site}>
                          {site}
                        </option>
                      ))}
                    </select>
                    <ChevronDown
                      size={18}
                      className="absolute right-4 top-1/2 -translate-y-1/2 text-muted pointer-events-none"
                    />
                  </div>
                </div>
              </div>

              {/* Dependent Toggle */}
              <div className="mt-6 pt-5 border-t border-divider">
                <label className="flex items-center gap-3 cursor-pointer group select-none">
                  <input
                    type="checkbox"
                    checked={wantDependent}
                    onChange={(e) => {
                      setWantDependent(e.target.checked);
                      if (e.target.checked) setActiveTab("associate");
                    }}
                    className="shrink-0"
                    id="dependent-toggle"
                  />
                  <UserPlus
                    size={20}
                    className={`transition-colors ${wantDependent ? "text-accent" : "text-muted"}`}
                  />
                  <span className="text-base font-medium text-foreground group-hover:text-accent transition-colors">
                    I also want to order for a dependent
                  </span>
                </label>
                {wantDependent && (
                  <p className="ml-12 mt-2 text-sm text-accent animate-fade-in flex items-center gap-1.5">
                    <Info size={14} />
                    Switch between baskets using the tabs below
                  </p>
                )}
              </div>
            </section>

            {/* ----------------------------------------------------------
                TABS (Associate / Dependent)
                ---------------------------------------------------------- */}
            <div className="flex gap-1 bg-muted-bg rounded-xl p-1">
              <button
                type="button"
                onClick={() => setActiveTab("associate")}
                className={`tab-button flex-1 flex items-center justify-center gap-2 py-3 px-4 rounded-lg text-base font-semibold transition-all ${activeTab === "associate"
                  ? "bg-card-bg text-primary shadow-sm active"
                  : "text-muted hover:text-foreground"
                  }`}
              >
                <User size={18} />
                My Basket
                {associateItemCount > 0 && (
                  <span className="bg-primary text-white text-xs font-bold px-2 py-0.5 rounded-full">
                    {associateItemCount}
                  </span>
                )}
              </button>
              {wantDependent && (
                <button
                  type="button"
                  onClick={() => setActiveTab("dependent")}
                  className={`tab-button flex-1 flex items-center justify-center gap-2 py-3 px-4 rounded-lg text-base font-semibold transition-all animate-fade-in ${activeTab === "dependent"
                    ? "bg-card-bg text-accent shadow-sm active"
                    : "text-muted hover:text-foreground"
                    }`}
                >
                  <Users size={18} />
                  Dependent&apos;s Basket
                  {dependentItemCount > 0 && (
                    <span className="bg-accent text-white text-xs font-bold px-2 py-0.5 rounded-full">
                      {dependentItemCount}
                    </span>
                  )}
                </button>
              )}
            </div>

            {/* ----------------------------------------------------------
                VACCINE GRID
                ---------------------------------------------------------- */}
            <section>
              {activeTab === "associate" ? (
                <div className="animate-fade-in">
                  <div className="flex items-center justify-between mb-4">
                    <h2 className="text-xl font-bold text-foreground flex items-center gap-2">
                      <ShoppingCart size={22} className="text-primary" />
                      Select Vaccines for Yourself
                    </h2>
                    {associateItemCount > 0 && (
                      <button
                        type="button"
                        onClick={() => setAssociateBasket(createEmptyBasket())}
                        className="text-sm text-danger hover:text-danger/80 transition-colors flex items-center gap-1"
                      >
                        <X size={14} /> Clear basket
                      </button>
                    )}
                  </div>
                  <BasketGrid
                    basket={associateBasket}
                    defaultSite={defaultSite}
                    onUpdate={setAssociateBasket}
                  />
                </div>
              ) : (
                <div className="animate-fade-in">
                  <div className="flex items-center justify-between mb-4">
                    <h2 className="text-xl font-bold text-foreground flex items-center gap-2">
                      <ShoppingCart size={22} className="text-accent" />
                      Select Vaccines for Dependent
                    </h2>
                    {dependentItemCount > 0 && (
                      <button
                        type="button"
                        onClick={() => setDependentBasket(createEmptyBasket())}
                        className="text-sm text-danger hover:text-danger/80 transition-colors flex items-center gap-1"
                      >
                        <X size={14} /> Clear basket
                      </button>
                    )}
                  </div>
                  <BasketGrid
                    basket={dependentBasket}
                    defaultSite={defaultSite}
                    onUpdate={setDependentBasket}
                  />
                </div>
              )}
            </section>

            {/* ----------------------------------------------------------
                ORDER SUMMARY
                ---------------------------------------------------------- */}
            <section
              className="rounded-2xl border-2 border-card-border bg-card-bg p-6 sm:p-8"
              style={{ boxShadow: "var(--shadow-sm)" }}
            >
              <h2 className="text-xl font-bold text-foreground mb-5 flex items-center gap-2">
                <Package size={22} className="text-primary" />
                Order Summary
              </h2>

              <OrderSummary
                associateBasket={associateBasket}
                dependentBasket={dependentBasket}
                wantDependent={wantDependent}
              />
            </section>

            {/* ----------------------------------------------------------
                SUBMIT RESULT MESSAGE
                ---------------------------------------------------------- */}
            {submitResult && (
              <div
                className={`rounded-2xl p-5 flex items-start gap-3 animate-fade-in ${submitResult.success
                  ? "bg-success-light border-2 border-success/30"
                  : "bg-danger-light border-2 border-danger/30"
                  }`}
              >
                {submitResult.success ? (
                  <CheckCircle2 size={24} className="text-success shrink-0 mt-0.5" />
                ) : (
                  <AlertCircle size={24} className="text-danger shrink-0 mt-0.5" />
                )}
                <div>
                  <p
                    className={`font-semibold text-base ${submitResult.success ? "text-success" : "text-danger"
                      }`}
                  >
                    {submitResult.success ? "Success!" : "Submission Failed"}
                  </p>
                  <p className="text-sm text-foreground mt-1">
                    {submitResult.message}
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => setSubmitResult(null)}
                  className="ml-auto shrink-0 text-muted hover:text-foreground transition-colors"
                  aria-label="Dismiss message"
                >
                  <X size={18} />
                </button>
              </div>
            )}

            {/* ----------------------------------------------------------
                SUBMIT BUTTON
                ---------------------------------------------------------- */}
            <div className="flex flex-col sm:flex-row items-center gap-4 pb-8">
              <button
                type="button"
                onClick={handleSubmit}
                disabled={!canSubmit || isSubmitting}
                className={`submit-btn flex-1 w-full sm:w-auto flex items-center justify-center gap-3 py-4 px-8 rounded-2xl text-lg font-bold text-white ${canSubmit && !isSubmitting
                  ? "bg-gradient-to-r from-primary to-primary-dark animate-pulse-soft cursor-pointer"
                  : "bg-muted cursor-not-allowed"
                  }`}
              >
                {isSubmitting ? (
                  <>
                    <Loader2 size={22} className="animate-spin" />
                    Submitting...
                  </>
                ) : (
                  <>
                    <Send size={22} />
                    Submit Order
                    {grandTotal > 0 && (
                      <span className="opacity-80">
                        ({formatPrice(grandTotal)})
                      </span>
                    )}
                  </>
                )}
              </button>

              {/* Validation hints */}
              {!canSubmit && !isSubmitting && (
                <div className="text-sm text-muted flex flex-col gap-1 animate-fade-in">
                  {!consentAgreed && (
                    <span className="flex items-center gap-1">
                      <AlertCircle size={14} className="text-warning" />
                      Accept the consent agreement
                    </span>
                  )}
                  {!associateNumber.trim() && (
                    <span className="flex items-center gap-1">
                      <AlertCircle size={14} className="text-warning" />
                      Enter your Associate Number
                    </span>
                  )}
                  {!defaultSite && (
                    <span className="flex items-center gap-1">
                      <AlertCircle size={14} className="text-warning" />
                      Select a default delivery site
                    </span>
                  )}
                  {totalItemCount === 0 && (
                    <span className="flex items-center gap-1">
                      <AlertCircle size={14} className="text-warning" />
                      Add at least one vaccine to your basket
                    </span>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    </main>
  );
}
