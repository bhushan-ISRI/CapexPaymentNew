import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState, useRef } from "react";
import { useNavigate } from "react-router-dom";
import logo from "../assets/sona-comstarlogo.png";
import "bootstrap/dist/css/bootstrap.min.css";
import Swal from "sweetalert2";

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

const EditAdvanceForm = ({ context, formData, onClose }: any) => {
  const sp = spfi().using(SPFx(context));
  const navigate = useNavigate();
  const today = new Date();
  const submitRef = useRef(false);
  const draftRef = useRef(false);
  const localDate: string = new Date(
    today.getTime() - today.getTimezoneOffset() * 60000,
  ).toISOString().split("T")[0];

  const [previousAdvances, setPreviousAdvances] = useState<IPreviousAdvance[]>([]);
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [employee, setEmployee] = useState<any>({});
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isDraftSaving, setIsDraftSaving] = useState(false);
  const [attachments, setAttachments] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
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
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);

  const handleNumberChange = (value: string, setter: any) => {
    const regex = /^\d*\.?\d*$/;
    if (regex.test(value)) setter(value);
  };

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

  const getVendors = async () => {
    try {
      const data = await sp.web.lists
        .getByTitle("VendorMaster")
        .items.select("Id", "VendorCode", "VendorName", "Status")
        .filter("Status eq 'Active'")();
      setVendors(data);
    } catch (error) {
      console.error("Vendor fetch error:", error);
    }
  };

  const getLoggedInUser = async () => {
    try {
      const currentUser = await sp.web.currentUser();
      const email = currentUser.Email;
      const user = await sp.web.lists
        .getByTitle("EmployeeMaster")
        .items.select(
          "EmployeeCode", "EmployeeName", "Division", "Location", "EmployeeEmail",
          "ReportingManager/Title", "ReportingManager/Id",
          "HOD/Title", "HOD/Id", "ContactNo", "EmployeeStatus", "CostCenter",
        )
        .expand("ReportingManager", "HOD")
        .filter(`EmployeeEmail eq '${email}'`)
        .top(1)();
      if (user.length > 0) setEmployee(user[0]);
    } catch (error) {
      console.log("Error fetching user:", error);
    }
  };

  const buildApprovalFlow = async () => {
    const flow: any[] = [];
    if (employee.ReportingManager?.Id)
      flow.push({ Id: employee.ReportingManager.Id, Name: employee.ReportingManager.Title, Role: "RM", Level: 1, Status: "Pending" });
    if (employee.HOD?.Id)
      flow.push({ Id: employee.HOD.Id, Name: employee.HOD.Title, Role: "HOD", Level: 2, Status: "Pending" });
    const matrixData = await sp.web.lists
      .getByTitle("CapexPaymentApprovalMatrix")
      .items.select("Role/RoleName,Approver/Id,Approver/Title,Level/Level")
      .expand("Approver,Role,Level")
      .filter("Status eq 'Active'")
      .orderBy("Level", true)();
    const matrixApprovers = matrixData.map((item: any, index: number) => ({
      Id: item.Approver?.Id,
      Name: item.Approver?.Title,
      Role: item.Role?.RoleName,
      Level: flow.length + index + 1,
      Status: "Pending",
    }));
    const finalFlow = [...flow, ...matrixApprovers];
    if (finalFlow.length > 0) finalFlow[0].Status = "In Progress";
    return finalFlow;
  };

  const getAttachments = async (capexId: string) => {
    try {
      if (!capexId) return;
      const safeCapexId = capexId.replace(/\//g, "_");
      const folderPath = `/sites/SonaFinance/CapexPaymentDocs/${safeCapexId}`;
      const files = await sp.web.getFolderByServerRelativePath(folderPath).files();
      setAttachments(files || []);
    } catch {
      setAttachments([]);
    }
  };

  const ensureFolder = async (folderPath: string) => {
    try {
      await sp.web.getFolderByServerRelativePath(folderPath)();
    } catch {
      const parentPath = folderPath.substring(0, folderPath.lastIndexOf("/"));
      const folderName = folderPath.substring(folderPath.lastIndexOf("/") + 1);
      await sp.web.getFolderByServerRelativePath(parentPath).folders.addUsingPath(folderName);
    }
  };

  const uploadFiles = async () => {
    try {
      if (!formData?.CapexId || selectedFiles.length === 0) return;
      const safe = formData.CapexId.replace(/\//g, "_");
      const folderPath = `/sites/SonaFinance/CapexPaymentDocs/${safe}`;
      await ensureFolder(folderPath);
      for (const file of selectedFiles) {
        await sp.web.getFolderByServerRelativePath(folderPath).files.addUsingPath(file.name, file, { Overwrite: true });
      }
      setSelectedFiles([]);
      await getAttachments(formData.CapexId);
    } catch (error) {
      console.error("Upload error:", error);
      await Swal.fire({ icon: "error", title: "Upload Failed", text: "File upload failed.", confirmButtonText: "OK" });
    }
  };

  const handleDeleteAttachment = async (fileName: string) => {
    try {
      if (!formData?.CapexId) return;
      const safeCapexId = formData.CapexId.replace(/\//g, "_");
      const folderPath = `/sites/SonaFinance/CapexPaymentDocs/${safeCapexId}`;
      await sp.web.getFolderByServerRelativePath(folderPath).files.getByUrl(fileName).recycle();
      setAttachments(attachments.filter((file: any) => file.Name !== fileName));
      await Swal.fire({ icon: "success", title: "Success", text: "Attachment deleted successfully.", confirmButtonText: "OK" });
    } catch (error) {
      console.error("Delete error:", error);
      await Swal.fire({ icon: "error", title: "Delete Failed", text: "Error deleting attachment.", confirmButtonText: "OK" });
    }
  };

  const validateForm = () => {
    const errors: string[] = [];
    if (!selectedVendorId || selectedVendorId === 0) errors.push("Please select Vendor");
    if (!poNumber || poNumber.trim() === "") errors.push("Please enter PO Number");
    if (!poDate || poDate === "Invalid Date") errors.push("Please select PO Date");
    if (poDate > localDate) errors.push("PO date cannot be a future date");
    if (!poTerms || poTerms.trim() === "") errors.push("Please enter PO Terms");
    if (!poAmount || poAmount.trim() === "") errors.push("Please enter PO Amount");
    if (!mrnNumber || mrnNumber.trim() === "") errors.push("Please enter MRN Number");
    if (!mrnDate || mrnDate === "Invalid Date") errors.push("Please select MRN Date");
    if (mrnDate > localDate) errors.push("MRN date cannot be a future date");
    if (!mrnAmount || mrnAmount.trim() === "") errors.push("Please enter MRN Amount");
    if (!requestedAmount || requestedAmount.trim() === "") errors.push("Please enter Requested Amount");
    if ((!attachments || attachments.length === 0) && (!selectedFiles || selectedFiles.length === 0))
      errors.push("Please upload at least one attachment");
    if (requestedAmount && mrnAmount && Number(requestedAmount) > Number(mrnAmount))
      errors.push("Requested Amount should not be greater than MRN Amount (GST)");
    return errors;
  };

  const handleSubmit = async () => {
    if (submitRef.current) return;
    setIsSubmitting(true);
    try {
      const errors = validateForm();
      if (errors.length > 0) {
        await Swal.fire({
          icon: "warning",
          title: "Validation",
          html: errors.map((err) => `• ${err}`).join("<br>"),
          confirmButtonText: "OK",
        });
        return;
      }
      const existingFlow = formData.ApprovalMatrix ? JSON.parse(formData.ApprovalMatrix) : [];
      const history = formData.WorkflowHistory ? JSON.parse(formData.WorkflowHistory) : [];
      history.push({
        CurrentApprover: employee.EmployeeName,
        ActionTaken: "Submitted",
        Comment: requesterRemarks,
        Date: new Date().toISOString(),
      });
      const currentApproverId = formData.CurrentApproverId || null;
      await sp.web.lists.getByTitle("CapexPayment").items.getById(formData.ID).update({
        Title: formData.CapexId,
        CapexId: formData.CapexId,
        EmployeeCode: employee.EmployeeCode,
        EmployeeName: employee.EmployeeName,
        Division: employee.Division,
        Location: employee.Location,
        Email: employee.EmployeeEmail,
        RM: employee.ReportingManager?.Title,
        HOD: employee.HOD?.Title,
        ContactNo: employee.ContactNo,
        EmployeeStatus: employee.EmployeeStatus,
        VendorCode: selectedVendorCode,
        VendorName: selectedVendorName,
        PONumber: poNumber,
        PODate: poDate ? new Date(poDate) : null,
        POPaymentTerms: poTerms,
        POAmount: poAmount ? poAmount.toString() : "",
        MRNNumber: mrnNumber,
        MRNDtae: mrnDate ? new Date(mrnDate) : null,
        MRNAmountwithGST: mrnAmount?.toString(),
        RequestedAmountforPayment: requestedAmount ? requestedAmount.toString() : "",
        FinalPaymentAgainstPO: finalPayment === "Yes",
        InstallationDetails: installationDetails, 
        RequesterRemarks: requesterRemarks,
        // Updated status: "Pending for Approval" on Submit
        StatusFlow: "Pending for Approval",
        Status: "Pending for Approval",
        ApprovalMatrix: JSON.stringify(existingFlow),
        CurrentApproverId: currentApproverId,
        WorkflowHistory: JSON.stringify(history),
      });
      if (selectedFiles.length > 0) await uploadFiles();
      await Swal.fire({ icon: "success", title: "Success", text: "Updated successfully.", confirmButtonText: "OK" });
      navigate("/User");
    } catch (error: any) {
      console.error("FULL ERROR:", error);
      await Swal.fire({ icon: "error", title: "Update Failed", text: error?.data?.responseBody || "Error while saving.", confirmButtonText: "OK" });
    } finally {
      submitRef.current = false;
      setIsSubmitting(false);
    }
  };

  const handledraft = async () => {
    if (isDraftSaving) return;
    setIsDraftSaving(true);
    try {
      const flow = await buildApprovalFlow();
      flow.forEach((f: any) => (f.Status = "Pending"));
      const currentApprover = flow.length > 0 ? flow[0].Id : null;
      const history = formData.WorkflowHistory ? JSON.parse(formData.WorkflowHistory) : [];
      history.push({
        CurrentApprover: employee.EmployeeName,
        ActionTaken: "Draft Saved",
        Comment: requesterRemarks || "",
        Date: new Date().toISOString(),
      });
      await sp.web.lists.getByTitle("CapexPayment").items.getById(formData.ID).update({
        Title: formData.CapexId,
        CapexId: formData.CapexId,
        EmployeeCode: employee.EmployeeCode,
        EmployeeName: employee.EmployeeName,
        Division: employee.Division,
        Location: employee.Location,
        Email: employee.EmployeeEmail,
        RM: employee.ReportingManager?.Title,
        HOD: employee.HOD?.Title,
        ContactNo: employee.ContactNo,
        EmployeeStatus: employee.EmployeeStatus,
        VendorCode: selectedVendorCode,
        VendorName: selectedVendorName,
        PONumber: poNumber,
        PODate: poDate ? new Date(poDate) : null,
        POPaymentTerms: poTerms,
        POAmount: poAmount ? poAmount.toString() : "",
        MRNNumber: mrnNumber,
        MRNDtae: mrnDate ? new Date(mrnDate) : null,
        MRNAmountwithGST: mrnAmount?.toString(),
        RequestedAmountforPayment: requestedAmount ? requestedAmount.toString() : "",
        FinalPaymentAgainstPO: finalPayment === "Yes",
        InstallationDetails: installationDetails,
        RequesterRemarks: requesterRemarks,
        // ✅ Status: "Draft" on Save as Draft
        StatusFlow: "Draft",
        Status: "Draft",
        ApprovalMatrix: JSON.stringify(flow),
        CurrentApproverId: currentApprover,
        WorkflowHistory: JSON.stringify(history),
      });
      if (selectedFiles.length > 0) await uploadFiles();
      await Swal.fire({ icon: "success", title: "Success", text: "Draft saved successfully.", confirmButtonText: "OK" });
      navigate("/User");
    } catch (error) {
      console.error("ERROR:", error);
      await Swal.fire({ icon: "error", title: "Save Failed", text: "Error while saving.", confirmButtonText: "OK" });
    } finally {
      draftRef.current = false;
      setIsDraftSaving(false);
    }
  };

  // ✅ Step 1: Bind all text fields from formData directly
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

  // ✅ Step 2: Once vendors load, match VendorCode string → Id for dropdown & load previous advances
  useEffect(() => {
    if (vendors.length > 0 && selectedVendorCode) {
      const match = vendors.find((v) => v.VendorCode === selectedVendorCode);
      if (match) {
        setSelectedVendorId(match.Id);
        setSelectedVendorName(match.VendorName);
        void getPreviousAdvances(match.Id);
      }
    }
  }, [vendors, selectedVendorCode]);

  useEffect(() => {
    void getLoggedInUser();
    void getVendors();
  }, []);

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            <div className="bordered">
              <img src={logo} />
              <h1>Edit Advance Payment</h1>
            </div>
            <ul className="approval-flow">
              <li className="approval-step">Initiator - {employee.EmployeeName}</li>
              {approvalMatrix.map((a, index) => (
                <li
                  key={index}
                  className={`approval-step ${
                    a.Status === "In Progress" ? "active" :
                    a.Status === "Approved" ? "approved" :
                    a.Status === "Rejected" ? "rejected" :
                    a.Status === "Send Back" ? "sendback" : ""
                  }`}
                >
                  {a.Role} - {a.Name}
                </li>
              ))}
            </ul>

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

              <div className="heading1" style={{ marginTop: "10px" }}><label>Vendor & PO Details</label></div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Vendor Code <span className="required">*</span></label>
                    <select
                      value={selectedVendorId || ""}
                      onChange={(e) => {
                        const id = Number(e.target.value);
                        const vendor = vendors.find((v) => v.Id === id);
                        setSelectedVendorId(id);
                        setSelectedVendorName(vendor?.VendorName || "");
                        setSelectedVendorCode(vendor?.VendorCode || "");
                        if (id > 0) void getPreviousAdvances(id);
                        else setPreviousAdvances([]);
                      }}
                      className="formtext-control"
                    >
                      <option value="">Select Vendor</option>
                      {vendors.map((v) => (
                        <option key={v.Id} value={v.Id}>{v.VendorCode}</option>
                      ))}
                    </select>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Name <span className="required">*</span></label>
                    <input value={selectedVendorName} className="form-control readonly" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number <span className="required">*</span></label>
                    <input value={poNumber} onChange={(e) => setPoNumber(e.target.value)} className="form-control" />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date <span className="required">*</span></label>
                    <input type="date" value={poDate} max={localDate} onChange={(e) => setPoDate(e.target.value)} className="form-control" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Advance Terms <span className="required">*</span></label>
                    <input value={poTerms} onChange={(e) => setPoTerms(e.target.value)} className="form-control" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount (GST) <span className="required">*</span></label>
                    <input value={poAmount} onChange={(e) => handleNumberChange(e.target.value, setPoAmount)} className="form-control" />
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}><label>MRN Details</label></div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">MRN Number <span className="required">*</span></label>
                    <input value={mrnNumber} onChange={(e) => setMrnNumber(e.target.value)} className="form-control" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Date <span className="required">*</span></label>
                    <input type="date" value={mrnDate} max={localDate} onChange={(e) => setMrnDate(e.target.value)} className="form-control" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Amount (GST) <span className="required">*</span></label>
                    <input value={mrnAmount} onChange={(e) => handleNumberChange(e.target.value, setMrnAmount)} className="form-control" />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Requested Amount for Payment <span className="required">*</span></label>
                    <input value={requestedAmount} onChange={(e) => handleNumberChange(e.target.value, setRequestedAmount)} className="form-control" />
                  </div>
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
                              <td colSpan={7} style={{ textAlign: "center", padding: "10px" }}>
                                {selectedVendorId ? "No previous advances available" : "Select a vendor to load advances"}
                              </td>
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

              <div className="heading1" style={{ marginTop: "10px" }}><label>Final Payment</label></div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Final Payment Against PO <span className="required">*</span></label>
                    <select value={finalPayment} onChange={(e) => setFinalPayment(e.target.value)} className="formtext-control">
                      <option value="No">No</option>
                      <option value="Yes">Yes</option>
                    </select>
                  </div>
                  {finalPayment === "Yes" && (
                    <div className="col-md-4 d-flex align-items-end">
                      <button
                        type="button"
                        onClick={() =>
                          window.open(
                            `${context.pageContext.web.absoluteUrl}/SitePages/Installation.aspx`,
                            "_blank",
                            "noopener,noreferrer",
                          )
                        }
                        style={{
                          backgroundColor: "#000",
                          color: "#fff",
                          padding: "6px 14px",
                          borderRadius: "4px",
                          border: "none",
                          fontSize: "14px",
                          cursor: "pointer",
                          display: "inline-block",
                        }}
                      >
                        Open Installation Form
                      </button>
                    </div>
                  )}
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}><label>Requester Remarks</label></div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-6">
                    <label className="font">Requester Remarks</label>
                    <textarea value={requesterRemarks} onChange={(e) => setRequesterRemarks(e.target.value)} className="form-control" rows={3} />
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}><label>Upload Document</label></div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attachments <span className="required" style={{ color: "red" }}>*</span></label>
                    {attachments.length > 0 && (
                      <ul className="mt-2">
                        {attachments.map((file: any, index: number) => (
                          <li key={index} style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "5px" }}>
                            <a href={file.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">{file.Name}</a>
                            <button type="button" className="btn btn-sm btn-danger" onClick={() => handleDeleteAttachment(file.Name)}>Delete</button>
                          </li>
                        ))}
                      </ul>
                    )}
                    {selectedFiles.length > 0 && (
                      <ul className="mt-2">
                        {selectedFiles.map((file: File, index: number) => (
                          <li key={index} style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "5px" }}>
                            <a href={URL.createObjectURL(file)} target="_blank" rel="noopener noreferrer">{file.name}</a>
                            <button type="button" className="btn btn-sm btn-danger" onClick={() => setSelectedFiles(selectedFiles.filter((_: File, i: number) => i !== index))}>Remove</button>
                          </li>
                        ))}
                      </ul>
                    )}
                    <input type="file" multiple className="form-control" onChange={(e) => { if (e.target.files) setSelectedFiles(Array.from(e.target.files)); }} />
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}><label>Workflow History</label></div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    {workflowHistory.length === 0 ? (
                      <p>No history available</p>
                    ) : (
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

              <div style={{ display: "flex", justifyContent: "center", gap: "10px", margin: "10px" }}>
                <button
                  type="button"
                  onClick={!isSubmitting ? handleSubmit : undefined}
                  disabled={isSubmitting}
                  className="submit-btn"
                  style={{ pointerEvents: isSubmitting ? "none" : "auto", opacity: isSubmitting ? 0.6 : 1 }}
                >
                  {isSubmitting ? "Submitting..." : "Submit"}
                </button>
                <button
                  type="button"
                  onClick={!isDraftSaving ? handledraft : undefined}
                  disabled={isDraftSaving}
                  className="Rework-btn"
                  style={{ pointerEvents: isDraftSaving ? "none" : "auto", opacity: isDraftSaving ? 0.6 : 1 }}
                >
                  {isDraftSaving ? "Saving..." : "Save as Draft"}
                </button>
                <a href="#" onClick={() => (onClose ? onClose() : navigate("/User"))} className="reset-btn">Exit</a>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default EditAdvanceForm;