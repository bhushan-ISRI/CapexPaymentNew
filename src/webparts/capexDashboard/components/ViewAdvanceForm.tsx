import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import 'bootstrap/dist/css/bootstrap.min.css';
import logo from "../assets/sona-comstarlogo.png";

interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

interface IPreviousAdvance {
  PONumber: string;
  RequestAdvanceAmount: string;
  Created: string;
  VoucherDate: string;
  VouchingNumber: string;
  PaidAmount: string;
  Status: string;
}

const ViewAdvanceForm = ({ context, formData, onClose }: any) => {
  const [attachments, setAttachments] = useState<any[]>([]);
  const sp = spfi().using(SPFx(context));
  const [employee, setEmployee] = useState<any>({});
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [previousAdvances, setPreviousAdvances] = useState<IPreviousAdvance[]>([]);
  const [mrnNumber, setMrnNumber] = useState("");
  const [mrnDate, setMrnDate] = useState("");
  const [mrnAmount, setMrnAmount] = useState("");
  const [requestedAmount, setRequestedAmount] = useState("");
  const [finalPayment, setFinalPayment] = useState("");
  const [installationDetails, setInstallationDetails] = useState("");
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [poAmount, setPoAmount] = useState("");
  const [requesterRemarks, setRequesterRemarks] = useState("");
  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [VouchingNumber, setVouchingNumber] = useState("");
  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);

  const getPreviousAdvances = async (vendorId: number) => {
    try {
      if (!vendorId) { setPreviousAdvances([]); return; }
      const data = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.select(
          "PONumber", "RequestAdvanceAmount", "Created",
          "VoucherDate", "VouchingNumber", "PaidAmount", "Status", "VendorCode/Id",
        )
        .expand("VendorCode")
        .filter(`VendorCode/Id eq ${vendorId} and Status eq 'Paid'`)
        .orderBy("Created", false)();
      setPreviousAdvances(data);
    } catch (error) {
      console.error("Error fetching previous advances:", error);
      setPreviousAdvances([]);
    }
  };

  const getAttachments = async (capexId: string) => {
    try {
      if (!capexId) return;
      const safeCapexId = capexId.replace(/\//g, "_");
      const folderPath = `/sites/SonaFinance/CapexPaymentDocs/${safeCapexId}`;
      const files = await sp.web.getFolderByServerRelativePath(folderPath).files();
      setAttachments(files || []);
    } catch { setAttachments([]); }
  };

  const getVendors = async () => {
    try {
      const data = await sp.web.lists.getByTitle("VendorMaster").items.select("Id", "VendorCode", "VendorName")();
      setVendors(data);
    } catch (error) { console.error("Vendor fetch error:", error); }
  };

  const getEmployeeDetails = async () => {
    try {
      if (!formData?.Email) return;
      const user = await sp.web.lists.getByTitle("EmployeeMaster").items
        .select("EmployeeCode", "EmployeeName", "Division", "Location", "EmployeeEmail",
          "ReportingManager/Title", "HOD/Title", "ContactNo", "EmployeeStatus", "CostCenter")
        .expand("ReportingManager", "HOD").filter(`EmployeeEmail eq '${formData.Email}'`).top(1)();
      if (user.length > 0) setEmployee(user[0]);
    } catch (error) { console.log("Error fetching employee:", error); }
  };

  useEffect(() => {
    if (!formData) return;
    setPoNumber(formData.PONumber || "");
    setPoDate(formData.PODate?.split("T")[0] || "");
    setPoTerms(formData.POPaymentTerms || "");
    setPoAmount(formData.POAmount || "");
    setSelectedVendorCode(formData.VendorCode || "");
    setSelectedVendorName(formData.VendorName || "");
    setMrnNumber(formData.MRNNumber || "");
    setMrnDate(formData.MRNDtae?.split("T")[0] || "");
    setMrnAmount(formData.MRNAmountwithGST || "");
    setRequestedAmount(formData.RequestedAmountforPayment || "");
    setFinalPayment(formData.FinalPaymentAgainstPO ? "Yes" : "No");
    setInstallationDetails(formData.InstallationDetails || "");
    setRequesterRemarks(formData.RequesterRemarks || "");
    setApproverRemarks(formData.ApproverRemarks || "");
    setVoucherDate(formData.VoucherDate?.split("T")[0] || "");
    setVouchingNumber(formData.VoucherNumber || "");
    setUTRDate(formData.UTRDate?.split("T")[0] || "");
    setUTRNumber(formData.UTRNumber || "");
    if (formData.CapexId) void getAttachments(formData.CapexId);
    if (formData?.ApprovalMatrix) {
      try {
        const parsed = typeof formData.ApprovalMatrix === "string" ? JSON.parse(formData.ApprovalMatrix) : formData.ApprovalMatrix;
        setApprovalMatrix(Array.isArray(parsed) ? parsed : []);
      } catch { setApprovalMatrix([]); }
    } else { setApprovalMatrix([]); }
    if (formData?.WorkflowHistory) {
      try {
        const parsed = typeof formData.WorkflowHistory === "string" ? JSON.parse(formData.WorkflowHistory) : formData.WorkflowHistory;
        setWorkflowHistory(Array.isArray(parsed) ? parsed : []);
      } catch { setWorkflowHistory([]); }
    } else { setWorkflowHistory([]); }
  }, [formData]);

  useEffect(() => {
    if (vendors.length > 0 && selectedVendorCode) {
      const match = vendors.find((v) => v.VendorCode === selectedVendorCode);
      if (match) {
        setSelectedVendorId(match.Id);
        void getPreviousAdvances(match.Id);
      }
    }
  }, [vendors, selectedVendorCode]);

  useEffect(() => { void getEmployeeDetails(); void getVendors(); }, []);

  const handleExit = () => { if (onClose) onClose(); else window.location.reload(); };

  // ✅ Helper: resolve step CSS class based on status
  const getStepClass = (status: string) => {
    switch (status) {
      case "In Progress": return "active";
      case "Approved": return "approved";
      case "Rejected": return "rejected";
      case "Send Back": return "sendback";
      case "Paid": return "approved";
      case "Pending for Vouching Update": return "active";
      case "Pending for UTR Update": return "active";
      default: return "";
    }
  };

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row"><div className="col-md-12"><div className="Main-Boxpoup">
        <div className="bordered"><img src={logo} /><h1>Advance Payment (View)</h1></div>

        {approvalMatrix.length === 0 ? (<p>No approval data</p>) : (
          <div className="displayWF">
            <ul className="approval-flow">
              {[
                { Role: "Initiator", Name: formData?.EmployeeName || employee.EmployeeName || "", Status: "Approved" },
                ...approvalMatrix.filter((a) => a.Role !== "Initiator")
              ].map((a, index) => (
                <li key={index} className={`approval-step ${getStepClass(a.Status)}`}>
                  {a.Role} - {a.Name}
                </li>
              ))}
            </ul>
          </div>
        )}

        <div className="borderedbox">
          <div className="heading1"><label>Requestor Information</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">Employee Code</label> : &nbsp;&nbsp;<label className="fonttext">{employee.EmployeeCode}</label></div>
              <div className="col-md-4"><label className="font">Employee Name</label> : &nbsp;&nbsp;<label className="fonttext">{employee.EmployeeName}</label></div>
              <div className="col-md-4"><label className="font">Employee Email</label> : &nbsp;&nbsp;<label className="fonttext">{employee.EmployeeEmail}</label></div>
            </div>
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">Contact No</label> : &nbsp;&nbsp;<label className="fonttext">{employee.ContactNo}</label></div>
              <div className="col-md-4"><label className="font">Employee Status</label> : &nbsp;&nbsp;<label className="fonttext">{employee.EmployeeStatus}</label></div>
              <div className="col-md-4"><label className="font">Division</label> : &nbsp;&nbsp;<label className="fonttext">{employee.Division}</label></div>
            </div>
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">Location</label> : &nbsp;&nbsp;<label className="fonttext">{employee.Location}</label></div>
              <div className="col-md-4"><label className="font">RM</label> : &nbsp;&nbsp;<label className="fonttext">{employee.ReportingManager?.Title}</label></div>
              <div className="col-md-4"><label className="font">HOD</label> : &nbsp;&nbsp;<label className="fonttext">{employee.HOD?.Title}</label></div>
            </div>
          </div>

          <div className="heading1" style={{ marginTop: "10px" }}><label>Vendor & PO</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">Vendor Code</label> : &nbsp;&nbsp;<label className="fonttext">{selectedVendorCode}</label></div>
              <div className="col-md-4"><label className="font">Vendor Name</label> : &nbsp;&nbsp;<label className="fonttext">{selectedVendorName}</label></div>
              <div className="col-md-4"><label className="font">PO Number</label> : &nbsp;&nbsp;<label className="fonttext">{poNumber}</label></div>
            </div>
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">PO Date</label> : &nbsp;&nbsp;<label className="fonttext">{poDate}</label></div>
              <div className="col-md-4"><label className="font">PO Amount</label> : &nbsp;&nbsp;<label className="fonttext">{poAmount}</label></div>
              <div className="col-md-4"><label className="font">PO Payment Terms</label> : &nbsp;&nbsp;<label className="fonttext">{poTerms}</label></div>
            </div>
          </div>

          <div className="heading1" style={{ marginTop: "10px" }}><label>MRN & Payment Details</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">MRN Number</label> : &nbsp;&nbsp;<label className="fonttext">{mrnNumber}</label></div>
              <div className="col-md-4"><label className="font">MRN Date</label> : &nbsp;&nbsp;<label className="fonttext">{mrnDate}</label></div>
              <div className="col-md-4"><label className="font">MRN Amount</label> : &nbsp;&nbsp;<label className="fonttext">{mrnAmount}</label></div>
            </div>
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">Requested Amount</label> : &nbsp;&nbsp;<label className="fonttext">{requestedAmount}</label></div>
            </div>
          </div>

          <div className="heading1" style={{ marginTop: "10px" }}><label>Previous Advances</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-12">
                <div style={{ overflowX: "auto" }}>
                  <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                    <thead className="text-white" style={{ backgroundColor: "rgb(60, 62, 69)" }}>
                      <tr>
                        <th className="px-4 py-2">PO Number</th>
                        <th className="px-4 py-2">Previous Advance</th>
                        <th className="px-4 py-2">Requested Date</th>
                        <th className="px-4 py-2">Paid Date</th>
                        <th className="px-4 py-2">Voucher No</th>
                        <th className="px-4 py-2">Settled Amount</th>
                        <th className="px-4 py-2">Pending Advance</th>
                      </tr>
                    </thead>
                    <tbody>
                      {previousAdvances.length === 0 ? (
                        <tr>
                          <td colSpan={7} style={{ textAlign: "center", padding: "10px" }}>No previous advances available</td>
                        </tr>
                      ) : (
                        previousAdvances.map((item: any, index: number) => {
                          const pending = Math.max(0, Number(item.RequestAdvanceAmount || 0) - Number(item.PaidAmount || 0));
                          return (
                            <tr key={index}>
                              <td className="px-4 py-2">{item.PONumber}</td>
                              <td className="px-4 py-2">{item.RequestAdvanceAmount}</td>
                              <td className="px-4 py-2">{item.Created ? new Date(item.Created).toLocaleDateString("en-GB") : ""}</td>
                              <td className="px-4 py-2">{item.VoucherDate ? new Date(item.VoucherDate).toLocaleDateString("en-GB") : ""}</td>
                              <td className="px-4 py-2">{item.VouchingNumber}</td>
                              <td className="px-4 py-2">{item.PaidAmount}</td>
                              <td className="px-4 py-2">{pending}</td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>

          <div className="heading1" style={{ marginTop: "10px" }}><label>Final Payment Details</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">Final Payment Against PO</label> : &nbsp;&nbsp;<label className="fonttext">{finalPayment}</label></div>
            </div>
            {finalPayment === "Yes" && (
              <div className="row mb-20">
                <div className="col-md-6"><label className="font">Installation Details</label> : &nbsp;&nbsp;<label className="fonttext">{installationDetails}</label></div>
              </div>
            )}
          </div>

          <div className="heading1" style={{ marginTop: "10px" }}><label>Requester Remarks</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-6"><label className="font">Requester Remarks</label> : &nbsp;&nbsp;<label className="fonttext">{requesterRemarks}</label></div>
            </div>
          </div>

          <div className="main-formcontainer" style={{ marginTop: "10px" }}>
            <div className="row mb-20">
              <div className="col-md-6"><label className="font">Voucher Date</label> : &nbsp;&nbsp;<label className="fonttext">{voucherDate}</label></div>
              <div className="col-md-6"><label className="font">Voucher Number</label> : &nbsp;&nbsp;<label className="fonttext">{VouchingNumber}</label></div>
            </div>
            <div className="row mb-20">
              <div className="col-md-6"><label className="font">UTR Date</label> : &nbsp;&nbsp;<label className="fonttext">{UTRDate}</label></div>
              <div className="col-md-6"><label className="font">UTR Number</label> : &nbsp;&nbsp;<label className="fonttext">{UTRNumber}</label></div>
            </div>
          </div>

          <div className="heading1" style={{ marginTop: "10px" }}><label>Upload Document</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-4">
                <label className="font">Attachments</label>
                {attachments.length > 0 ? (
                  <ul>{attachments.map((file: any, index: number) => (
                    <li key={index}><a href={file.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">{file.Name}</a></li>
                  ))}</ul>
                ) : <p>No attachments</p>}
              </div>
            </div>
          </div>

          <div className="heading1" style={{ marginTop: "10px" }}><label>Workflow History</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-12">
                {workflowHistory.length === 0 ? (<p>No history available</p>) : (
                  <table className="workflow-table" style={{ width: "100%" }}>
                    <thead>
                      <tr>
                        <th style={{ padding: "8px", textAlign: "left" }}>Action By</th>
                        <th style={{ padding: "8px", textAlign: "left" }}>Action Taken</th>
                        <th style={{ padding: "8px", textAlign: "left" }}>Date</th>
                        <th style={{ padding: "8px", textAlign: "left" }}>Comment</th>
                      </tr>
                    </thead>
                    <tbody>
                      {workflowHistory
                        .filter((h: any) => h.ActionTaken && h.ActionTaken !== "Edited")
                        .map((h: any, idx: number) => (
                          <tr key={idx}>
                            <td style={{ padding: "8px" }}>{h.CurrentApprover || ""}</td>
                            <td style={{ padding: "8px" }}>{h.ActionTaken || ""}</td>
                            <td style={{ padding: "8px" }}>{h.Date ? new Date(h.Date).toLocaleDateString("en-GB") : ""}</td>
                            <td style={{ padding: "8px" }}>{h.Comment || ""}</td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                )}
              </div>
            </div>
          </div>

          <div className="row my-3">
            <div className="col-md-12">
              <div className="text-center">
                <a href="#" onClick={handleExit} className="reset-btn">Exit</a>
              </div>
            </div>
          </div>
        </div>
      </div></div></div>
    </div>
  );
};

export default ViewAdvanceForm;