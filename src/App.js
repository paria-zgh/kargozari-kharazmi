import React, { useState } from "react";
import Main from "./Main";
import Enerzhi from "./Enerzhi";
import Amar from "./Amar";
import AmarEnergy from "./AmarEnergy";
import "./App.css";
import logo from "./Assets/logo.jpg";
import Footer from "./Footer";
import Smain from "./Smain";

import Siman from "./Siman";
import { Petro } from "./Petro";
import MazadEnerji from "./MazadEnerji";

export default function App() {
  const [activeMarket, setActiveMarket] = useState(null);
  const [mode, setMode] = useState(null);
  const [excelType, setExcelType] = useState(null);
  const resetAll = () => {
    setActiveMarket(null);
    setMode(null);
    setExcelType(null);
  };

  const handleSelectMarket = (market) => {
    setActiveMarket(market);
    setMode(null);
    setExcelType(null);
  };

  return (
    <div className="app-root" dir="rtl">

      {/* ===== Header ===== */}
      <header className="app-header">
        <div className="header-inline">
          <img src={logo} className="app-logo" alt="لوگو" />
          <div>
            <h2>کارگزاری آینده نگر خوارزمی</h2>
            <p className="header-subtitle">سامانه پردازش داده</p>
          </div>
        </div>
      </header>

      <div className="container main-container">

        {/* ===== انتخاب بازار ===== */}
        {!activeMarket && (
          <div className="row justify-content-center g-4 mb-4 main-buttons-row">
            <div className="col-md-5">
              <div
                className="select-main-card"
                onClick={() => handleSelectMarket("kala")}
              >
                <h3>بورس کالا</h3>
              </div>
            </div>

            <div className="col-md-5">
              <div
                className="select-main-card energy-main"
                onClick={() => handleSelectMarket("energy")}>
                <h3>بورس انرژی</h3>
              </div>
            </div>
          </div>
        )}

        {/* ===== عنوان بازار ===== */}
        {activeMarket && (
          <div className="top-title-box">
            <h2>{activeMarket === "kala" ? "بورس کالا" : "بورس انرژی"}</h2>
          </div>
        )}

        {/* ===== بورس کالا ===== */}
        {activeMarket === "kala" && (
          <>
            <div className="row justify-content-center g-3 mb-4 sub-buttons-row">

              <div className="col-md-3 col-6">
                <button
                  className={`sub-btn ${mode === "Smain" ? "active" : ""}`}
                  onClick={() => {
                    setMode("Smain");
                    setExcelType(null);
                  }}
                >
                  (سیمرغ)عرضه‌های بورس کالا
                </button>
              </div>

              <div className="col-md-3 col-6">
                <button
                  className={`sub-btn ${mode === "amar" ? "active" : ""}`}
                  onClick={() => {
                    setMode("amar");
                    setExcelType(null);
                  }}
                >
                  آمار بورس کالا
                </button>
              </div>
              <div className="col-md-3 col-6">
                <button
                  className={`sub-btn ${mode === "amar" ? "active" : ""}`}
                  onClick={() => {
                    setMode("amar");
                    setExcelType(null);
                  }}
                >
                  آمار بورس کالا
                </button>
              </div>
              <div className="col-md-3 col-6">
                <button
                  className={`sub-btn ${mode === "excel" ? "active" : ""}`}
                  onClick={() => {
                    setMode("excel");
                    setExcelType(null);
                  }}
                >
                  اکسل بورس کالا
                </button>
              </div>

            </div>

            {/* زیرمنوی اکسل */}
            {mode === "excel" && (
              <div className="row justify-content-center g-3 mb-4">
                <div className="col-md-3 col-6">
                  <button
                    className={`sub-btn ${excelType === "petro" ? "active" : ""}`}
                    onClick={() => setExcelType("petro")}
                  >
                    اکسل پتروشیمی
                  </button>
                </div>

                <div className="col-md-3 col-6">
                  <button
                    className={`sub-btn ${excelType === "siman" ? "active" : ""}`}
                    onClick={() => setExcelType("siman")}
                  >
                    اکسل سیمان
                  </button>
                </div>
              </div>
            )}

            {(mode || excelType) && (
              <div className="content-card">
                {mode === "main" && <Main />}
                {mode === "Smain" && <Smain />}
                {mode === "amar" && <Amar />}
                {excelType === "petro" && <Petro />}
                {excelType === "siman" && <Siman />}
                
              </div>
            )}

            {/* دکمه بازگشت پایین صفحه */}
            <div className="back-bottom">
              <button className="sub-btn back-btn" onClick={resetAll}>
                بازگشت
              </button>
            </div>
          </>
        )}

        {/* ===== بورس انرژی ===== */}
        {activeMarket === "energy" && (
          <>
            <div className="row justify-content-center g-3 mb-4 sub-buttons-row">

              <div className="col-md-3 col-6">
                <button
                  className={`sub-btn ${mode === "energyMain" ? "active" : ""}`}
                  onClick={() => {
                    setMode("energyMain");
                    setExcelType(null);
                  }}
                >
                  عرضه‌های بورس انرژی
                </button>
              </div>

              <div className="col-md-3 col-6">
                <button
                  className={`sub-btn ${mode === "energyAmar" ? "active" : ""}`}
                  onClick={() => {
                    setMode("energyAmar");
                    setExcelType(null);
                  }}
                >
                  آمار بورس انرژی
                </button>
              </div>
              <div className="col-md-3 col-6">
                <button
                  className={`sub-btn ${mode === "mazadEnerji" ? "active" : ""}`}
                  onClick={() => {
                    setMode("mazadEnerji");
                    setExcelType(null);
                  }}
                >
                  عرضه مازاد انرژی
                </button>
              </div>
            </div>

            {mode && (
              <div className="content-card">
                {mode === "energyMain" && <Enerzhi />}
                {mode === "energyAmar" && <AmarEnergy />}
                {mode === "mazadEnerji" && <MazadEnerji />}
              </div>
            )}

            {/* دکمه بازگشت پایین صفحه */}
            <div className="back-bottom">
              <button className="sub-btn back-btn" onClick={resetAll}>
                بازگشت
              </button>
            </div>
          </>
        )}

      </div>

    </div>
  );
}
