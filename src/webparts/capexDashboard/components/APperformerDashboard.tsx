import * as React from "react";
import "./userDashboardsc.scss";
import APperformerAdvanceFormForUTR from "./APperformerAdvanceFormForUTR";

import { useState } from "react";
import APperformerAdvanceform from "./APperformerAdvanceform";
import logo from "../assets/SonaPNGLogo.png";
import Edit from "../assets/Pencil.png";
import User from "../assets/Userlogo.png";
import View from "../assets/Eye.png";
import ViewAdvanceForm from "./ViewAdvanceForm";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

interface UserDashboardProps {
  context: any;
}

const APperformerDashboard: React.FC<UserDashboardProps> = ({ context }) => {
  const sp = spfi().using(SPFx(context));
  const [formType, setFormType] = useState<
    "approve" | "approveUTR" | "view" | null
  >(null);
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 10;

  const [activeMenu, setActiveMenu] = React.useState("My Request");
  const [searchText, setSearchText] = React.useState("");
  const [statusFilter, setStatusFilter] = React.useState("");
  const [showForm, setShowForm] = React.useState(false);
  const [data, setData] = React.useState<any[]>([]);
  const [currentUserName, setCurrentUserName] = React.useState("");
  const [selectedItem, setSelectedItem] = React.useState<any>(null);

  const handleFormOpen = async (
    item: any,
    type: "view" | "approve" | "approveUTR",
  ) => {
    try {
      debugger;
      const fullItem = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(item.ID)
        .select(
          "*",
          "ID",
          "Title",
          "EmployeeCode",
          "EmployeeName",
          "Email",
          "Division",
          "Location",
          "ContactNo",
          "EmployeeStatus",
          "RM",
          "HOD",
          "CapexId",
          "VendorName",
          "VendorCode",
          "PONumber",
          "PODate",
          "POAmount",
          "POPaymentTerms",
          "MRNNumber",
          "MRNDtae",
          "MRNAmountwithGST",
          "RequestedAmountforPayment",
          "VoucherNumber",
          "VoucherDate",
          "ApproverRemarks",
          "FinalPaymentAgainstPO",
          "InstallationDetails",
          "Status",
          "Author/EMail",
        )
        .expand("Author")();

      const mappedData = {
        ...fullItem,
        EmployeeEmail: fullItem.EmployeeEmail,
        POAmtGST: fullItem.POAmount,
        POAdvanceTerms: fullItem.POPaymentTerms,
        mrnNumber: fullItem.MRNNumber,
        mrnDate: fullItem.MRNDtae,
        mrnAmount: fullItem.MRNAmountwithGST,
        requestedAmount: fullItem.RequestedAmountforPayment,
        finalPayment: fullItem.FinalPaymentAgainstPO,
        installationDetails: fullItem.InstallationDetails,
        ReportingManager: fullItem.ReportingManager,
        HOD: fullItem.HOD,
        ApprovalMatrix: fullItem.ApprovalMatrix
          ? JSON.parse(fullItem.ApprovalMatrix)
          : null,
        WorkflowHistory: fullItem.WorkflowHistory
          ? JSON.parse(fullItem.WorkflowHistory)
          : null,
      };

      setSelectedItem(mappedData);
      setFormType(type);
      setShowForm(true);
    } catch (error) {
      console.error("Form open error:", error);
    }
  };

  const getLoggedInUser = async () => {
    try {
      const user = await sp.web.currentUser();
      setCurrentUserName(user.Title);
    } catch (error) {
      console.error("User error:", error);
    }
  };

  const handleApproveClick = async (item: any) => {
    debugger;
    try {
      const fullItem = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(item.ID)
        .select("*", "Author/EMail")
        .expand("Author")();

      setSelectedItem(fullItem);

      if (item.status === "Pending for Vouching Update") {
        setFormType("approve");
      } else if (item.status === "Pending for UTR Update") {
        setFormType("approveUTR");
      }

      setShowForm(true);
    } catch (error) {
      console.error("View error:", error);
    }
  };

  const handleViewClick = async (item: any) => {
    try {
      const fullItem = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(item.ID)
        .select("*", "Author/EMail")
        .expand("Author")();

      setSelectedItem(fullItem);
      setFormType("view");
      setShowForm(true);
    } catch (error) {
      console.error("View error:", error);
    }
  };

  const getCapexData = async () => {
    try {
      const currentUser = await sp.web.currentUser();
      let filterQuery;

      if (activeMenu === "My Request") {
        filterQuery = `
    (
      Status eq 'Pending for Vouching Update'
      or Status eq 'Pending for UTR Update'
    )
    and CurrentApproverId eq ${currentUser.Id}
  `
          .replace(/\n/g, "")
          .trim();
      } else if (activeMenu === "Paid") {
        filterQuery = `
    Status eq 'Paid'
    and Author/Id eq ${currentUser.Id}
  `
          .replace(/\n/g, "")
          .trim();
      } else if (activeMenu === "Rejected") {
        filterQuery = `
    Status eq 'Reject'
    and Author/Id eq ${currentUser.Id}
  `
          .replace(/\n/g, "")
          .trim();
      }

      const items = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.select(
          "*",
          "ID",
          "Title",
          "Created",
          "EmployeeName",
          "VendorName",
          "VendorCode",
          "PONumber",
          "RequestedAmountforPayment",
          "Status",
          "CurrentApproverId",
          "CurrentApprover/Title",
          "CurrentApprover/EMail",
          "PendingWth",
          "Author/EMail",
        )
        .expand("Author", "CurrentApprover")
        .filter(filterQuery)
        .orderBy("ID", false)();

      const formatted = items.map((item: any) => ({
        ID: item.ID,
        id: item.Title,
        date: item.Created
          ? new Date(item.Created).toLocaleDateString("en-GB")
          : "",
        EmployeeName: item.EmployeeName || "",
        vendor: item.VendorName || "",
        vendorCode: item.VendorCode || "",
        po: item.PONumber || "",
        amount: item.RequestedAmountforPayment || 0,
        pendingWith: item.CurrentApprover?.Title || item.PendingWth || "",
        status: item.Status || "",
      }));

      setData(formatted);
    } catch (error) {
      console.error("Data error:", error);
    }
  };

  const filteredData = data.filter((item) => {
    const text = searchText.toLowerCase();
    const status = statusFilter.toLowerCase();

    let menuFilter = true;
    if (activeMenu === "Paid") {
      menuFilter = item.status?.toLowerCase() === "paid";
    } else if (activeMenu === "Rejected") {
      menuFilter = item.status?.toLowerCase() === "reject";
    } else if (activeMenu === "My Request") {
      menuFilter = true;
    }

    return (
      menuFilter &&
      (item.id?.toLowerCase().includes(text) ||
        item.vendor?.toLowerCase().includes(text) ||
        item.po?.toLowerCase().includes(text)) &&
      (!status || item.status?.toLowerCase().includes(status))
    );
  });

  const totalPages = Math.ceil(filteredData.length / itemsPerPage);
  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const paginatedData = filteredData.slice(startIndex, endIndex);

  React.useEffect(() => {
    if (!context) return;
    void getLoggedInUser();
    void getCapexData();
  }, [context, activeMenu]);

  if (showForm) {
    if (formType === "approve") {
      return (
        <APperformerAdvanceform context={context} itemId={selectedItem?.ID} />
      );
    }
    if (formType === "approveUTR") {
      return (
        <APperformerAdvanceFormForUTR
          context={context}
          itemId={selectedItem?.ID}
        />
      );
    }
    if (formType === "view") {
      return (
        <ViewAdvanceForm
          context={context}
          formData={selectedItem}
          onClose={() => {
            setShowForm(false);
            setFormType(null);
            void getCapexData();
          }}
        />
      );
    }
  }

  return (
    <div style={{ display: "flex", width: "100%" }}>
      <div className="sidebar">
        <div className="sidehead">
          <div className="logo">
            <img src={logo} width="25px" height="25px" />
          </div>
          <div className="sidehead-right">SONA COMSTAR</div>
        </div>

        <div className="sidehead-user">
          <img
            src={User}
            style={{ margin: "10px 20px" }}
            width={20}
            height={20}
          />
          {currentUserName}
        </div>

        <ul className="nav">
          <li className="nav-item">
            <a
              className={
                activeMenu === "My Request" ? "nav-link active" : "nav-link"
              }
              onClick={() => setActiveMenu("My Request")}
              style={{ cursor: "pointer" }}
            >
              My Request
            </a>
          </li>
          <li className="nav-item">
            <a
              className={activeMenu === "Paid" ? "nav-link active" : "nav-link"}
              onClick={() => setActiveMenu("Paid")}
              style={{ cursor: "pointer" }}
            >
              Paid
            </a>
          </li>
          <li className="nav-item">
            <a
              className={
                activeMenu === "Rejected" ? "nav-link active" : "nav-link"
              }
              onClick={() => setActiveMenu("Rejected")}
              style={{ cursor: "pointer" }}
            >
              Rejected
            </a>
          </li>
        </ul>
      </div>

      <div
        className="main"
        style={{ width: "calc(100% - 250px)", transition: "width 0.3s" }}
      >
        <div className="header">
          <div className="left-banner">
            <div className="logo-text">
              <h2>Capex Payment Performer Dashboard</h2>
            </div>
          </div>
        </div>

        <div className="col-md-12 mainsecond">
          <div>
            <input
              placeholder="Search"
              value={searchText}
              className="form-control"
              style={{ width: "250px" }}
              onChange={(e) => setSearchText(e.target.value)}
            />
          </div>
          <div>
            <select
              value={statusFilter}
              className="formtext-control"
              onChange={(e) => setStatusFilter(e.target.value)}
            >
              <option value="">All</option>
              <option value="Pending for Approval">Pending for Approval</option>
              <option value="Approved">Approved</option>
              <option value="Rejected">Rejected</option>
              <option value="Send Back">Send Back</option>
              <option value="Draft">Draft</option>
              <option value="Pending for Vouching Update">
                Pending for Vouching Update
              </option>
              <option value="Paid">Paid</option>
              <option value="Pending for UTR Update">
                Pending for UTR Update
              </option>
            </select>
          </div>
        </div>

        <main className="Main-Dash mx-2">
          <div style={{ overflowX: "auto" }}>
            <div className="table-vert-scroll">
              <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                <thead
                  className="text-white"
                  style={{ backgroundColor: "rgb(60, 62, 69)" }}
                >
                  <tr>
                    <th className="px-4 py-2">Action</th>
                    <th className="px-4 py-2">Payment ID</th>
                    <th className="px-4 py-2">Requestor Date</th>
                    <th className="px-4 py-2">Requestor Name</th>
                    <th className="px-4 py-2">Requestor Type</th>
                    <th className="px-4 py-2">Vendor Code</th>
                    <th className="px-4 py-2">Vendor Name</th>
                    <th className="px-4 py-2">PO Number</th>
                    <th className="px-4 py-2">Request Amount</th>
                    <th className="px-4 py-2">Pending With</th>
                    <th className="px-4 py-2">Status</th>
                  </tr>
                </thead>
                <tbody>
                  {paginatedData.length === 0 ? (
                    <tr>
                      <td colSpan={11} style={{ textAlign: "center" }}>
                        No Data
                      </td>
                    </tr>
                  ) : (
                    paginatedData.map((item, i) => (
                      <tr key={i}>
                        <td className="px-4 py-2">
                          {(activeMenu === "Paid" ||
                            activeMenu === "Rejected") && (
                            <span
                              onClick={() => handleViewClick(item)}
                              style={{ cursor: "pointer", marginRight: "10px" }}
                            >
                              <img src={View} width={15} alt="View" />
                            </span>
                          )}
                          {activeMenu === "My Request" &&
                            (item.status === "Pending for Vouching Update" ||
                              item.status === "Pending for UTR Update") && (
                              <span
                                onClick={() => handleApproveClick(item)}
                                style={{ cursor: "pointer" }}
                              >
                                <img src={Edit} width={15} alt="Edit" />
                              </span>
                            )}
                        </td>
                        <td className="px-4 py-2">{item.id}</td>
                        <td className="px-4 py-2">{item.date}</td>
                        <td className="px-4 py-2">{item.EmployeeName}</td>
                        <td className="px-4 py-2">Capex Payment</td>
                        <td className="px-4 py-2">{item.vendorCode}</td>
                        <td className="px-4 py-2">{item.vendor}</td>
                        <td className="px-4 py-2">{item.po}</td>
                        <td className="px-4 py-2">₹ {item.amount}</td>
                        <td className="px-4 py-2">{item.pendingWith || "-"}</td>
                        <td className="px-4 py-2">{item.status}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>

              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  gap: "10px",
                  marginTop: "15px",
                }}
              >
                <button
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage(currentPage - 1)}
                >
                  Previous
                </button>
                <span>
                  Page {currentPage} of {totalPages || 1}
                </span>
                <button
                  disabled={currentPage === totalPages || totalPages === 0}
                  onClick={() => setCurrentPage(currentPage + 1)}
                >
                  Next
                </button>
              </div>
            </div>
          </div>
        </main>
      </div>
    </div>
  );
};

export default APperformerDashboard;
