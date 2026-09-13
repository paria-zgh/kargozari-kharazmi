import React from "react";
import "./Footer.css";
export default function Footer() {
  return (
    <footer className="app-footer">
      <div className="footer-container">

        <div className="footer-item">
          <strong><i className="fas fa-phone-alt"></i> شماره تماس:</strong>
          <p><i className="fas fa-mobile-alt"></i> 09123011311</p>
          <p><i className="fas fa-mobile-alt"></i> 09001221476</p>
        </div>

        <div className="footer-item">
          <strong><i className="fas fa-phone"></i> تلفن ثابت:</strong>
          <p><i className="fas fa-phone"></i> 02161914401</p>
        </div>

        <div className="footer-item">
          <strong><i className="fas fa-envelope"></i> ایمیل:</strong>
          <p><i className="fas fa-envelope"></i> <a href="mailto:paria.zoghi73@gmail.com" style={{ color: "#fff" }}>paria.zoghi73@gmail.com</a></p>
        </div>

        <div className="footer-item">
          <strong><i className="fas fa-map-marker-alt"></i> آدرس:</strong>
          <p>
            <i className="fas fa-map-marker-alt"></i> ولیعصر، خیابان بزرگمهر، مجتمع تجاری اداری بزرگمهر، پلاک 16، واحد 101
          </p>
        </div>

      </div>
    </footer>
  );
}
