import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState, useRef } from "react";
import { useNavigate } from "react-router-dom";
import "bootstrap/dist/css/bootstrap.min.css";
import logo from "../assets/sona-comstarlogo.png";
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

const NewAdvanceform = ({ context, onClose }: any) => {
  const navigate = useNavigate();
  const submitRef = useRef(false);
  const draftRef = useRef(false);
  const sp = spfi().using(SPFx(context));
  const today = new Date();
  const localDate: string = new Date(
    today.getTime() - today.getTimezoneOffset() * 60000,
  )
    .toISOString()
    .split("T")[0];

  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isDraftSaving, setIsDraftSaving] = useState(false);
  const [employee, setEmployee] = React.useState<any>({});
  const [attachments, setAttachments] = useState<File[]>([]);
  const [previousAdvances, setPreviousAdvances] = useState<IPreviousAdvance[]>(
    [],
  );
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [finalPayment, setFinalPayment] = useState("");
  const [installationDetails, setInstallationDetails] = useState("");
  const [poAmount, setPoAmount] = useState("");
  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [mrnNumber, setMrnNumber] = useState("");
  const [mrnDate, setMrnDate] = useState("");
  const [mrnAmount, setMrnAmount] = useState("");
  const [requestedAmount, setRequestedAmount] = useState("");
  const [requesterRemarks, setRequesterRemarks] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);

  const handleNumberChange = (value: string, setter: any) => {
    const regex = /^\d*\.?\d*$/;
    if (regex.test(value)) setter(value);
  };

  const handleRemoveFile = (index: number) => {
    const updatedFiles = [...attachments];
    updatedFiles.splice(index, 1);
    setAttachments(updatedFiles);
  };

  const getPreviousAdvances = async (vendorId: number) => {
    try {
      if (!vendorId) {
        setPreviousAdvances([]);
        return;
      }
      const data = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.select(
          "PONumber",
          "RequestAdvanceAmount",
          "Created",
          "VoucherDate",
          "VouchingNumber",
          "PaidAmount",
          "Status",
          "VendorCode/Id",
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

  const getLoggedInUser = async () => {
    try {
      const currentUser = await sp.web.currentUser();
      const email = currentUser.Email;
      const user = await sp.web.lists
        .getByTitle("EmployeeMaster")
        .items.select(
          "EmployeeCode",
          "EmployeeName",
          "Division",
          "Location",
          "EmployeeEmail",
          "ReportingManager/Title",
          "ReportingManager/Id",
          "HOD/Title",
          "HOD/Id",
          "ContactNo",
          "EmployeeStatus",
          "CostCenter",
        )
        .expand("ReportingManager", "HOD")
        .filter(`EmployeeEmail eq '${email}'`)
        .top(1)();
      if (user.length > 0) setEmployee(user[0]);
      buildApprovalPreview(user[0]);
    } catch (error) {
      console.log("Error fetching user:", error);
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

  const getFinancialYear = () => {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;
    if (month >= 4)
      return `${year.toString().slice(-2)}-${(year + 1).toString().slice(-2)}`;
    return `${(year - 1).toString().slice(-2)}-${year.toString().slice(-2)}`;
  };

  const generateCapexId = async () => {
    try {
      const fy = getFinancialYear();
      const items = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.select("CapexId", "ID")
        .filter(`startswith(CapexId,'CPX-PYMT/${fy}/')`)
        .orderBy("ID", false)
        .top(1)();
      let nextNumber = 1;
      if (items.length > 0 && items[0].CapexId) {
        const parts = items[0].CapexId.split("/");
        if (parts.length === 3) {
          const lastNumber = parseInt(parts[2]);
          if (!isNaN(lastNumber)) nextNumber = lastNumber + 1;
        }
      }
      return `CPX-PYMT/${fy}/${nextNumber.toString().padStart(5, "0")}`;
    } catch (error) {
      return `CPX-PYMT/${getFinancialYear()}/00001`;
    }
  };

  const uploadAttachments = async (capexId: string) => {
    try {
      if (!attachments || attachments.length === 0) return;
      const safeCapexId = capexId.replace(/\//g, "_");
      const libraryName = "CapexPaymentDocs";
      const webUrl = context.pageContext.web.serverRelativeUrl;
      const folderPath = `${webUrl}/${libraryName}/${safeCapexId}`;
      await sp.web.folders.addUsingPath(`${libraryName}/${safeCapexId}`);
      for (const file of attachments) {
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, file, { Overwrite: true });
      }
    } catch (error) {
      console.error("Upload error:", error);
    }
  };

  const buildApprovalPreview = async (emp: any) => {
    const flow: any[] = [];
    if (emp.ReportingManager?.Title)
      flow.push({
        Name: emp.ReportingManager.Title,
        Role: "RM",
        Status: "Pending",
      });
    if (emp.HOD?.Title)
      flow.push({ Name: emp.HOD.Title, Role: "HOD", Status: "Pending" });
    const matrixData = await sp.web.lists
      .getByTitle("CapexPaymentApprovalMatrix")
      .items.select("Role/RoleName,Approver/Title")
      .expand("Approver,Role")
      .filter("Status eq 'Active'")
      .orderBy("Level", true)();
    const matrixApprovers = matrixData.map((item: any) => ({
      Name: item.Approver?.Title,
      Role: item.Role?.RoleName,
      Status: "Pending",
    }));
    setApprovalMatrix([...flow, ...matrixApprovers]);
  };

  const buildApprovalFlow = async () => {
    const flow: any[] = [];
    if (employee.ReportingManager?.Id)
      flow.push({
        Id: employee.ReportingManager.Id,
        Name: employee.ReportingManager.Title,
        Role: "RM",
        Level: 1,
        Status: "Pending",
      });
    if (employee.HOD?.Id)
      flow.push({
        Id: employee.HOD.Id,
        Name: employee.HOD.Title,
        Role: "HOD",
        Level: 2,
        Status: "Pending",
      });
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

  const validateForm = () => {
    const errors: string[] = [];
    if (!selectedVendorId || selectedVendorId === 0)
      errors.push("Please select Vendor");
    if (!poNumber || poNumber.trim() === "")
      errors.push("Please enter PO Number");
    if (!poDate || poDate === "Invalid Date")
      errors.push("Please select PO Date");
    if (poDate > localDate) errors.push("PO date cannot be a future date");
    if (!poTerms || poTerms.trim() === "") errors.push("Please enter PO Terms");
    if (!poAmount || poAmount.trim() === "")
      errors.push("Please enter PO Amount");
    if (!mrnNumber || mrnNumber.trim() === "")
      errors.push("Please enter MRN Number");
    if (!mrnDate || mrnDate === "Invalid Date")
      errors.push("Please select MRN Date");
    if (mrnDate > localDate) errors.push("MRN date cannot be a future date");
    if (!mrnAmount || mrnAmount.trim() === "")
      errors.push("Please enter MRN Amount");
    if (!requestedAmount || requestedAmount.trim() === "")
      errors.push("Please enter Requested Amount");
    if (finalPayment === "Yes" && !installationDetails)
      errors.push("Please enter Installation Details");
    if (!attachments || attachments.length === 0)
      errors.push("Please upload attachment");
    if (
      requestedAmount &&
      mrnAmount &&
      Number(requestedAmount) > Number(mrnAmount)
    )
      errors.push(
        "Requested Amount should not be greater than MRN Amount (GST)",
      );
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
      const capexId = await generateCapexId();
      const flow = await buildApprovalFlow();
      const currentApprover = flow.length > 0 ? flow[0].Id : null;
      const WorkflowHistory = [
        {
          CurrentApprover: employee.EmployeeName,
          ActionTaken: "Submitted",
          Comment: requesterRemarks,
          Date: new Date().toISOString(),
        },
      ];
      await sp.web.lists.getByTitle("CapexPayment").items.add({
        Title: capexId,
        CapexId: capexId,
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
        RequestedAmountforPayment: requestedAmount
          ? requestedAmount.toString()
          : "",
        FinalPaymentAgainstPO: finalPayment === "Yes",
        InstallationDetails: installationDetails,
        RequesterRemarks: requesterRemarks,
        StatusFlow: "Pending for Approval",
        Status: "Pending for Approval",
        ApprovalMatrix: JSON.stringify(flow),
        CurrentApproverId: currentApprover,
        WorkflowHistory: JSON.stringify(WorkflowHistory),
      });
      await uploadAttachments(capexId);
      await Swal.fire({
        icon: "success",
        title: "Success",
        text: "Submitted successfully.",
        confirmButtonText: "OK",
      });
      navigate("/User");
    } catch (error: any) {
      console.error("FULL ERROR:", error);
      await Swal.fire({
        icon: "error",
        title: "Submission Failed",
        text: error?.data?.responseBody || "Error while saving.",
        confirmButtonText: "OK",
      });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handledraft = async () => {
    if (draftRef.current) return;
    setIsDraftSaving(true);
    try {
      const capexId = await generateCapexId();
      const flow = await buildApprovalFlow();
      flow.forEach((f: any) => (f.Status = "Pending"));
      const currentApprover = flow.length > 0 ? flow[0].Id : null;
      const WorkflowHistory = [
        {
          CurrentApprover: employee.EmployeeName,
          ActionTaken: "Draft Saved",
          Comment: requesterRemarks || "",
          Date: new Date().toISOString(),
        },
      ];
      await sp.web.lists.getByTitle("CapexPayment").items.add({
        Title: capexId,
        CapexId: capexId,
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
        RequestedAmountforPayment: requestedAmount
          ? requestedAmount.toString()
          : "",
        FinalPaymentAgainstPO: finalPayment === "Yes",
        InstallationDetails: installationDetails,
        RequesterRemarks: requesterRemarks,
        StatusFlow: "Draft",
        Status: "Draft",
        ApprovalMatrix: JSON.stringify(flow),
        CurrentApproverId: currentApprover,
        WorkflowHistory: JSON.stringify(WorkflowHistory),
      });
      void uploadAttachments(capexId);
      await Swal.fire({
        icon: "success",
        title: "Success",
        text: "Draft saved successfully.",
        confirmButtonText: "OK",
      });
      navigate("/User");
    } catch (error) {
      console.error("ERROR:", error);
      await Swal.fire({
        icon: "error",
        title: "Save Failed",
        text: "Error while saving.",
        confirmButtonText: "OK",
      });
    } finally {
      draftRef.current = false;
      setIsDraftSaving(false);
    }
  };

  React.useEffect(() => {
    if (!context) return;
    void getLoggedInUser();
    void getVendors();
  }, [context]);

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            <div className="bordered">
              <img src={logo} />
              <h1>Capex Payment Request</h1>
            </div>

            {approvalMatrix.length === 0 ? (
              <p>Loading...</p>
            ) : (
              <div className="displayWF">
                <ul className="approval-flow">
                  <li className="approval-step">
                    Initiator - {employee.EmployeeName}
                  </li>
                  {approvalMatrix.map((a, index) => (
                    <li
                      key={index}
                      className={`approval-step ${index === 0 ? "active" : ""}`}
                    >
                      {a.Role} - {a.Name}
                    </li>
                  ))}
                </ul>
              </div>
            )}

            <div className="borderedbox">
              <div className="heading1">
                <label>Requestor Information</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Employee Code</label> : &nbsp;&nbsp;
                    <label className="fonttext">{employee.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Name</label> : &nbsp;&nbsp;
                    <label className="fonttext">{employee.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Email</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">{employee.EmployeeEmail}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Contact No</label> : &nbsp;&nbsp;
                    <label className="fonttext">{employee.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Status</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">
                      {employee.EmployeeStatus}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Division</label> : &nbsp;&nbsp;
                    <label className="fonttext">{employee.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Location</label> : &nbsp;&nbsp;
                    <label className="fonttext">{employee.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">RM</label> : &nbsp;&nbsp;
                    <label className="fonttext">
                      {employee.ReportingManager?.Title}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">HOD</label> : &nbsp;&nbsp;
                    <label className="fonttext">{employee.HOD?.Title}</label>
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Vendor & PO Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">
                      Vendor Code&nbsp;<span className="required">*</span>
                    </label>
                    <select
                      value={selectedVendorId || ""}
                      onChange={(e) => {
                        const id = Number(e.target.value);
                        const vendor = vendors.find((v) => v.Id === id);
                        setSelectedVendorId(id);
                        setSelectedVendorName(vendor?.VendorName || "");
                        setSelectedVendorCode(vendor?.VendorCode || "");
                        setPoNumber("");
                        if (id > 0) void getPreviousAdvances(id);
                        else setPreviousAdvances([]);
                      }}
                      className="formtext-control"
                    >
                      <option value="">Select Vendor</option>
                      {vendors.map((v) => (
                        <option key={v.Id} value={v.Id}>
                          {v.VendorCode}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Name</label>
                    <input
                      value={selectedVendorName}
                      className="form-control readonly"
                      readOnly
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">
                      PO Number&nbsp;<span className="required">*</span>
                    </label>
                    <input
                      value={poNumber}
                      className="form-control"
                      onChange={(e) => setPoNumber(e.target.value)}
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">
                      PO Date&nbsp;<span className="required">*</span>
                    </label>
                    <input
                      type="date"
                      value={poDate}
                      className="form-control"
                      max={localDate}
                      onChange={(e) => setPoDate(e.target.value)}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">
                      PO Amount (GST)&nbsp;<span className="required">*</span>
                    </label>
                    <input
                      value={poAmount}
                      className="form-control"
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setPoAmount)
                      }
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">
                      PO Payment Terms&nbsp;<span className="required">*</span>
                    </label>
                    <input
                      value={poTerms}
                      className="form-control"
                      onChange={(e) => setPoTerms(e.target.value)}
                    />
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Payment Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">
                      MRN Number&nbsp;<span className="required">*</span>
                    </label>
                    <input
                      value={mrnNumber}
                      onChange={(e) => setMrnNumber(e.target.value)}
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">
                      MRN Date&nbsp;<span className="required">*</span>
                    </label>
                    <input
                      type="date"
                      value={mrnDate}
                      max={localDate}
                      onChange={(e) => setMrnDate(e.target.value)}
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">
                      MRN Amount Including (GST)&nbsp;
                      <span className="required">*</span>
                    </label>
                    <input
                      value={mrnAmount}
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setMrnAmount)
                      }
                      className="form-control"
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">
                      Requested Amount for Payment&nbsp;
                      <span className="required">*</span>
                    </label>
                    <input
                      value={requestedAmount}
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setRequestedAmount)
                      }
                      className="form-control"
                    />
                  </div>
                </div>
              </div>

              <div className="main-formcontainer" style={{ marginTop: "10px" }}>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font fontblock">
                      Whether this is the Final Payment against the PO&nbsp;
                      <span className="required">*</span>
                    </label>
                    <div className="radioalign">
                      <label className="fonttext">
                        <div>
                          <input
                            type="radio"
                            value="Yes"
                            checked={finalPayment === "Yes"}
                            onChange={(e) => setFinalPayment(e.target.value)}
                          />
                        </div>
                        &nbsp;<div>Yes</div>
                      </label>
                      &nbsp;&nbsp;
                      <label className="fonttext">
                        <div>
                          <input
                            type="radio"
                            value="No"
                            checked={finalPayment === "No"}
                            onChange={(e) => setFinalPayment(e.target.value)}
                          />
                        </div>
                        &nbsp;<div>No</div>
                      </label>
                    </div>
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

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Previous Advances</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    <div style={{ overflowX: "auto" }}>
                      <div className="table-vert-scroll">
                        <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                          <thead
                            className="text-white"
                            style={{ backgroundColor: "rgb(60, 62, 69)" }}
                          >
                            <tr>
                              <th className="px-4 py-2">PO Number</th>
                              <th className="px-4 py-2">Previous Advance</th>
                              <th className="px-4 py-2">Requested Date</th>
                              <th className="px-4 py-2">Paid Date</th>
                              <th className="px-4 py-2">MRN No</th>
                              <th className="px-4 py-2">Settled Amount</th>
                              <th className="px-4 py-2">Pending Advance</th>
                            </tr>
                          </thead>
                          <tbody>
                            {previousAdvances.length === 0 ? (
                              <tr>
                                <td colSpan={7} style={{ textAlign: "center" }}>
                                  No previous advances available
                                </td>
                              </tr>
                            ) : (
                              previousAdvances.map(
                                (item: any, index: number) => {
                                  const pending = Math.max(
                                    0,
                                    Number(item.RequestAdvanceAmount || 0) -
                                      Number(item.PaidAmount || 0),
                                  );
                                  return (
                                    <tr key={index}>
                                      <td>{item.PONumber}</td>
                                      <td>{item.RequestAdvanceAmount}</td>
                                      <td>
                                        {item.Created
                                          ? new Date(
                                              item.Created,
                                            ).toLocaleDateString()
                                          : ""}
                                      </td>
                                      <td>
                                        {item.VoucherDate
                                          ? new Date(
                                              item.VoucherDate,
                                            ).toLocaleDateString()
                                          : ""}
                                      </td>
                                      <td>{item.VoucherNumber}</td>
                                      <td>{item.PaidAmount}</td>
                                      <td>{pending}</td>
                                    </tr>
                                  );
                                },
                              )
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Requester Remarks</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-6">
                    <textarea
                      value={requesterRemarks}
                      onChange={(e) => setRequesterRemarks(e.target.value)}
                      className="form-control"
                      rows={5}
                      style={{ height: "100px" }}
                      placeholder="Enter remarks..."
                    />
                  </div>
                </div>
              </div>

              <div className="main-formcontainer" style={{ marginTop: "10px" }}>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">
                      Attach{" "}
                      <span className="required" style={{ color: "red" }}>
                        *
                      </span>
                    </label>
                    <input
                      type="file"
                      multiple
                      className="form-control"
                      onChange={(e) => {
                        if (e.target.files)
                          setAttachments((prev) => [
                            ...prev,
                            ...Array.from(e.target.files!),
                          ]);
                      }}
                    />
                    {attachments.length > 0 && (
                      <ul style={{ marginTop: "10px" }}>
                        {attachments.map((file: File, index: number) => (
                          <li key={index}>
                            <a
                              href={URL.createObjectURL(file)}
                              target="_blank"
                              rel="noopener noreferrer"
                            >
                              {file.name}
                            </a>
                            <button
                              type="button"
                              style={{
                                marginLeft: "10px",
                                color: "red",
                                cursor: "pointer",
                              }}
                              onClick={() => handleRemoveFile(index)}
                            >
                              Remove
                            </button>
                          </li>
                        ))}
                      </ul>
                    )}
                  </div>
                </div>
              </div>

              <div className="row my-3">
                <div className="col-md-12">
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "center",
                      alignItems: "center",
                      gap: "10px",
                    }}
                  >
                    <button
                      type="button"
                      onClick={!isSubmitting ? handleSubmit : undefined}
                      disabled={isSubmitting}
                      className="submit-btn"
                      style={{
                        pointerEvents: isSubmitting ? "none" : "auto",
                        opacity: isSubmitting ? 0.6 : 1,
                      }}
                    >
                      {isSubmitting ? "Submitting..." : "Submit"}
                    </button>
                    <button
                      type="button"
                      onClick={!isDraftSaving ? handledraft : undefined}
                      disabled={isDraftSaving}
                      className="Rework-btn"
                      style={{
                        pointerEvents: isDraftSaving ? "none" : "auto",
                        opacity: isDraftSaving ? 0.6 : 1,
                      }}
                    >
                      {isDraftSaving ? "Saving..." : "Save as Draft"}
                    </button>
                    <a
                      href="#"
                      onClick={() => navigate("/User")}
                      className="reset-btn"
                    >
                      Exit
                    </a>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default NewAdvanceform;