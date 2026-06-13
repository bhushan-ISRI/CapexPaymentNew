import * as React from "react";
import { HashRouter, Routes, Route, useNavigate } from "react-router-dom";
import { ICapexDashboardProps } from "./ICapexDashboardProps";
import UserDashboard from "./UserDashboard";
import ApproverDashboard from "./ApproverDashboard";
import APperformerDashboard from "./APperformerDashboard";
import ApLogo from "../assets/ApDashboard.png";
import UserLogo from "../assets/UserDashboard.png";
import ApproverLogo from "../assets/ApproverDashboard.png";
import "./usersite.scss";
import "../assets/bootstrap/css/bootstrap.css";

function HomePage(props: ICapexDashboardProps) {
  const navigate = useNavigate();

  return (
    <div className="main-container">
      <div className="headSheet">
        <h2>Capex Payment</h2>
      </div>
      <section className="hero">
        <div className="overlay"></div>
        <div className="hero-content">
          <div className="card-container">
            <div className="infoCard" onClick={() => navigate("/User")}>
              <div className="cardContent">
                <div className="cardalin">
                  <span className="boximage">
                    <img src={UserLogo} width="25" height="25" />
                  </span>
                  <h4>User Dashboard</h4>
                </div>
              </div>
            </div>

            <div className="infoCard" onClick={() => navigate("/Approver")}>
              <div className="cardContent">
                <div className="cardalin">
                  <span className="boximage">
                    <img src={ApproverLogo} width="25" height="25" />
                  </span>
                  <h4>Approver Dashboard</h4>
                </div>
              </div>
            </div>

            <div className="infoCard" onClick={() => navigate("/Performer")}>
              <div className="cardContent">
                <div className="cardalin">
                  <span className="boximage">
                    <img src={ApLogo} width="25" height="25" />
                  </span>
                  <h4>Ap Performer Dashboard</h4>
                </div>
              </div>
            </div>
          </div>
        </div>
      </section>
    </div>
  );
}

export default function CapexDashboard(props: ICapexDashboardProps) {
  return (
    <HashRouter>
      <Routes>
        <Route path="/" element={<HomePage {...props} />} />
        <Route path="/User" element={<UserDashboard {...props} />} />
        <Route path="/Approver" element={<ApproverDashboard {...props} />} />
        <Route
          path="/Performer"
          element={<APperformerDashboard {...props} />}
        />
      </Routes>
    </HashRouter>
  );
}
