import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState, useRef } from "react";
import logo from "../assets/sona-comstarlogo.png";
import "bootstrap/dist/css/bootstrap.min.css";
import Swal from "sweetalert2";
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';

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
  const [mrnBasicAmount, setMrnBasicAmount] = useState("");
  const [mrnGstAmount, setMrnGstAmount] = useState("");
  const [mrnOtherAmount, setMrnOtherAmount] = useState("");
  const [requestedAmount, setRequestedAmount] = useState("");
  const [finalPayment, setFinalPayment] = useState("");
  const [installationDetails, setInstallationDetails] = useState("");
  const [installationRequestNumber, setInstallationRequestNumber] = useState("");
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [poBasicAmount, setPoBasicAmount] = useState("");
  const [poGstAmount, setPoGstAmount] = useState("");
  const [poOtherAmount, setPoOtherAmount] = useState("");
  const [requesterRemarks, setRequesterRemarks] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);

  const [vendorSearch, setVendorSearch] = useState("");
  const [vendorDropdownOpen, setVendorDropdownOpen] = useState(false);
  const vendorDropdownRef = useRef<HTMLDivElement>(null);
  const vendorSearchRef = useRef<HTMLInputElement>(null);
  const [pageLoading, setPageLoading] = useState(true);

  const filteredVendors = vendors.filter(
    (v) =>
      v.VendorName.toLowerCase().includes(vendorSearch.toLowerCase()) ||
      v.VendorCode.toLowerCase().includes(vendorSearch.toLowerCase()),
  );

  const totalPoAmount =
    (Number(poBasicAmount) || 0) +
    (Number(poGstAmount) || 0) +
    (Number(poOtherAmount) || 0);

  const totalMrnAmount =
    (Number(mrnBasicAmount) || 0) +
    (Number(mrnGstAmount) || 0) +
    (Number(mrnOtherAmount) || 0);

  const toDouble = (val: string): number | null =>
    val && val.trim() !== "" ? parseFloat(val) : null;

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

  // const getLoggedInUser = async () => {
  //   try {
  //     const currentUser = await sp.web.currentUser();
  //     const email = currentUser.Email;
  //     const user = await sp.web.lists
  //       .getByTitle("EmployeeMaster")
  //       .items.select(
  //         "EmployeeCode", "EmployeeName", "Division", "Location", "EmployeeEmail",
  //         "ReportingManager/Title", "ReportingManager/Id",
  //         "HOD/Title", "HOD/Id", "ContactNo", "EmployeeStatus", "CostCenter",
  //       )
  //       .expand("ReportingManager", "HOD")
  //       .filter(`EmployeeEmail eq '${email}'`)
  //       .top(1)();
  //     if (user.length > 0) setEmployee(user[0]);
  //   } catch (error) {
  //     console.log("Error fetching user:", error);
  //   }
  // };
  const ensureUser = async (email: string): Promise<number> => {

    if (!email) return 0;

    try {

      const webUrl = context.pageContext.web.absoluteUrl;

      const response = await context.spHttpClient.post(
        `${webUrl}/_api/web/ensureuser`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Accept": "application/json;odata=nometadata",
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            logonName: email
          })
        }
      );

      if (!response.ok) {

        console.log("ensureUser failed for:", email);

        return 0;
      }

      const data = await response.json();

      return data.Id || 0;

    } catch (error) {

      console.log("ensureUser error:", email, error);

      return 0;
    }
  };
  const getLoggedInUser = async () => {
    try {
      const toTitleCase = (str: string): string => {
        if (!str) return "";

        return str
          .toLowerCase()
          .split(" ")
          .filter(Boolean)
          .map((w: string) => w.charAt(0).toUpperCase() + w.slice(1))
          .join(" ");
      };

      const cleanLocationForDisplay = (location: string): string => {
        if (!location) return "";
        return location.replace(/^re\s+/i, "").trim();
      };

      const FLOW_URL =
        "https://defaultcb1edbfe8080457d9cae51528f3643.3f.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/e2bb522aa41443179a72b701b9613471/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=q8b8ADCtK2eKr2f6p3MX7gxmJymPeJbm0mq2M69Rk8E";

      const fetchPage = async (pageNumber: number) => {
        const response = await fetch(FLOW_URL, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            PageSize: 500,
            PageNumber: pageNumber,
          }),
        });

        if (!response.ok) {
          throw new Error("Failed to fetch employee data");
        }

        return response.json();
      };

      const currentUserEmail =
        context.pageContext.user.email.toLowerCase();

      let employee: any = null;
      let page = 1;

      while (true) {
        const res = await fetchPage(page);

        const employees = res?.data?.employees || [];

        employee = employees.find(
          (x: any) => x.email?.toLowerCase() === currentUserEmail
        );

        if (employee) break;

        if (employees.length < 500) break;

        page++;
      }

      if (!employee) {
        console.log("Employee not found.");
        return;
      }

      const attributes = employee.attributes || [];

      const locationAttr = attributes.find(
        (x: any) =>
          x.attributeTypeDescription?.toLowerCase() === "location"
      );

      const departmentAttr = attributes.find(
        (x: any) =>
          x.attributeTypeDescription?.toLowerCase() === "department"
      );

      const hodEmailAttr = attributes.find(
        (x: any) =>
          x.attributeTypeDescription?.toLowerCase() === "hod_email"
      );

      const hodNameAttr = attributes.find(
        (x: any) =>
          x.attributeTypeDescription?.toLowerCase() === "hod name"
      );

      let rmUserId = 0;
      let hodUserId = 0;

      try {
        if (employee.reportingManagerEmail) {
          rmUserId = await ensureUser(employee.reportingManagerEmail);
        }

        if (hodEmailAttr?.attributeTypeUnitDescription) {
          hodUserId = await ensureUser(
            hodEmailAttr.attributeTypeUnitDescription
          );
        }
      } catch (err) {
        console.log("ensureUser error:", err);
      }



      setEmployee({
        EmployeeCode: employee.employeeCode || "",
        EmployeeName: toTitleCase(employee.employeeName || ""),
        Division: departmentAttr?.attributeTypeUnitDescription || "",
        Location: cleanLocationForDisplay(
          locationAttr?.attributeTypeUnitDescription || ""
        ),
        EmployeeEmail: employee.email || "",
        ContactNo: employee.mobileNo || "",
        EmployeeStatus: employee.employeeStatus || "",
        CostCenter: employee.costCenter || "",

        ReportingManager: {
          Id: rmUserId,
          Title: employee.reportingManagerName || "",
        },

        HOD: {
          Id: hodUserId,
          Title: hodNameAttr?.attributeTypeUnitDescription || "",
        },
      });

    } catch (error) {
      console.error("Error fetching user:", error);
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
    if (!poBasicAmount || poBasicAmount.trim() === "") errors.push("Please enter PO Basic Amount");
    if (!poGstAmount || poGstAmount.trim() === "") errors.push("Please enter PO GST Amount");
    if (!poOtherAmount || poOtherAmount.trim() === "") errors.push("Please enter PO Other Amount");
    if (!mrnNumber || mrnNumber.trim() === "") errors.push("Please enter MRN Number");
    if (!mrnDate || mrnDate === "Invalid Date") errors.push("Please select MRN Date");
    if (mrnDate > localDate) errors.push("MRN date cannot be a future date");
    if (!mrnBasicAmount || mrnBasicAmount.trim() === "") errors.push("Please enter MRN Basic Amount");
    if (!mrnGstAmount || mrnGstAmount.trim() === "") errors.push("Please enter MRN GST Amount");
    if (!mrnOtherAmount || mrnOtherAmount.trim() === "") errors.push("Please enter MRN Other Amount");
    if (!requestedAmount || requestedAmount.trim() === "") errors.push("Please enter Requested Amount");
    if (
      finalPayment === "No" &&
      (!installationRequestNumber || installationRequestNumber.trim() === "")
    ) errors.push("Please enter Installation Request Number");
    if ((!attachments || attachments.length === 0) && (!selectedFiles || selectedFiles.length === 0))
      errors.push("Please upload at least one attachment");
    if (requestedAmount && totalMrnAmount && Number(requestedAmount) > totalMrnAmount)
      errors.push("Requested Amount should not be greater than MRN Amount (GST)");
    return errors;
  };

  const buildPayload = (flow: any[], status: string, history: any[], isDraft: boolean) => {
    const currentApproverId = flow.length > 0 ? flow[0].Id : null;
    const poTotal = totalPoAmount || null;
    const mrnTotal = totalMrnAmount || null;
    return {
      Title: formData.CapexId,
      CapexId: formData.CapexId,
      EmployeeCode: employee.EmployeeCode || null,
      EmployeeName: employee.EmployeeName || null,
      Division: employee.Division || null,
      Location: employee.Location || null,
      Email: employee.EmployeeEmail || null,
      RM: employee.ReportingManager?.Title || null,
      HOD: employee.HOD?.Title || null,
      ContactNo: employee.ContactNo || null,
      EmployeeStatus: employee.EmployeeStatus || null,
      VendorCode: selectedVendorCode || null,
      VendorName: selectedVendorName || null,
      PONumber: poNumber || null,
      PODate: poDate ? new Date(poDate) : null,
      POPaymentTerms: poTerms || null,
      POBasicAmount: isDraft ? toDouble(poBasicAmount) : parseFloat(poBasicAmount),
      POGSTAmount: isDraft ? toDouble(poGstAmount) : parseFloat(poGstAmount),
      POOtherAmount: isDraft ? toDouble(poOtherAmount) : parseFloat(poOtherAmount),
      POAmount: poTotal !== null ? poTotal.toString() : null,
      MRNNumber: mrnNumber || null,
      MRNDtae: mrnDate ? new Date(mrnDate) : null,
      MRNBasicAmount: isDraft ? toDouble(mrnBasicAmount) : parseFloat(mrnBasicAmount),
      MRNGSTAmount: isDraft ? toDouble(mrnGstAmount) : parseFloat(mrnGstAmount),
      MRNOtherAmount: isDraft ? toDouble(mrnOtherAmount) : parseFloat(mrnOtherAmount),
      MRNAmountwithGST: mrnTotal !== null ? mrnTotal.toString() : null,
      RequestedAmountforPayment: requestedAmount ? requestedAmount.toString() : null,
      FinalPaymentAgainstPO: finalPayment === "Yes" ? true : finalPayment === "No" ? false : null,
      InstallationDetails: installationDetails || null,
      InstallationRequestNumber: finalPayment === "No" ? installationRequestNumber || null : null,
      RequesterRemarks: requesterRemarks || null,
      StatusFlow: status,
      Status: status,
      ApprovalMatrix: JSON.stringify(flow),
      CurrentApproverId: currentApproverId,
      WorkflowHistory: JSON.stringify(history),
    };
  };

  const handleSubmit = async () => {
    if (submitRef.current) return;
    submitRef.current = true;
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
      const history = formData.WorkflowHistory ? JSON.parse(formData.WorkflowHistory) : [];
      history.push({
        CurrentApprover: employee.EmployeeName,
        ActionTaken: "Submitted",
        Comment: requesterRemarks,
        Date: new Date().toISOString(),
      });
      const flow = await buildApprovalFlow();
      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(formData.ID)
        .update(buildPayload(flow, "Pending for Approval", history, false));
      if (selectedFiles.length > 0) await uploadFiles();
      await Swal.fire({ icon: "success", title: "Success", text: "Updated successfully.", confirmButtonText: "OK" });
      onClose();
    } catch (error: any) {
      console.error("FULL ERROR:", error);
      await Swal.fire({
        icon: "error",
        title: "Update Failed",
        text: error?.data?.responseBody || error?.message || "Error while saving.",
        confirmButtonText: "OK",
      });
    } finally {
      submitRef.current = false;
      setIsSubmitting(false);
    }
  };

  const handledraft = async () => {
    if (draftRef.current) return;
    draftRef.current = true;
    setIsDraftSaving(true);
    try {
      const flow = await buildApprovalFlow();
      flow.forEach((f: any) => (f.Status = "Pending"));
      const history = formData.WorkflowHistory ? JSON.parse(formData.WorkflowHistory) : [];
      history.push({
        CurrentApprover: employee.EmployeeName,
        Comment: requesterRemarks || "",
        Date: new Date().toISOString(),
      });
      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(formData.ID)
        .update(buildPayload(flow, "Draft", history, true));
      if (selectedFiles.length > 0) await uploadFiles();
      await Swal.fire({ icon: "success", title: "Success", text: "Draft saved successfully.", confirmButtonText: "OK" });
      onClose();
    } catch (error: any) {
      console.error("ERROR:", error);
      await Swal.fire({
        icon: "error",
        title: "Save Failed",
        text: error?.data?.responseBody || error?.message || "Error while saving.",
        confirmButtonText: "OK",
      });
    } finally {
      draftRef.current = false;
      setIsDraftSaving(false);
    }
  };

  useEffect(() => {
    if (!formData) return;
    setPoNumber(formData.PONumber || "");
    setPoDate(formData.PODate?.split("T")[0] || "");
    setPoTerms(formData.POPaymentTerms || "");
    setPoBasicAmount(formData.POBasicAmount != null ? String(formData.POBasicAmount) : "");
    setPoGstAmount(formData.POGSTAmount != null ? String(formData.POGSTAmount) : "");
    setPoOtherAmount(formData.POOtherAmount != null ? String(formData.POOtherAmount) : "");
    setSelectedVendorCode(formData.VendorCode || "");
    setSelectedVendorName(formData.VendorName || "");
    setMrnNumber(formData.MRNNumber || "");
    setMrnDate(formData.MRNDtae?.split("T")[0] || "");
    setMrnBasicAmount(formData.MRNBasicAmount != null ? String(formData.MRNBasicAmount) : "");
    setMrnGstAmount(formData.MRNGSTAmount != null ? String(formData.MRNGSTAmount) : "");
    setMrnOtherAmount(formData.MRNOtherAmount != null ? String(formData.MRNOtherAmount) : "");
    setRequestedAmount(formData.RequestedAmountforPayment || "");
    setFinalPayment(formData.FinalPaymentAgainstPO ? "Yes" : "No");
    setInstallationDetails(formData.InstallationDetails || "");
    setInstallationRequestNumber(formData.InstallationRequestNumber || "");
    setRequesterRemarks(formData.RequesterRemarks || "");
    if (formData.CapexId) void getAttachments(formData.CapexId);
    if (formData?.ApprovalMatrix) {
      try {
        const parsed = typeof formData.ApprovalMatrix === "string"
          ? JSON.parse(formData.ApprovalMatrix)
          : formData.ApprovalMatrix;
        setApprovalMatrix(Array.isArray(parsed) ? parsed : []);
      } catch { setApprovalMatrix([]); }
    } else { setApprovalMatrix([]); }
    if (formData?.WorkflowHistory) {
      try {
        const parsed = typeof formData.WorkflowHistory === "string"
          ? JSON.parse(formData.WorkflowHistory)
          : formData.WorkflowHistory;
        setWorkflowHistory(Array.isArray(parsed) ? parsed : []);
      } catch { setWorkflowHistory([]); }
    } else { setWorkflowHistory([]); }
  }, [formData]);

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
  void loadPage();
}, []);

 const loadPage = async () => {
  try {
    setPageLoading(true);

    await Promise.all([
      getLoggedInUser(),
      getVendors(),
    ]);
  } catch (err) {
    console.error(err);
  } finally {
    setPageLoading(false);
  }
};

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (
        vendorDropdownRef.current &&
        !vendorDropdownRef.current.contains(event.target as Node)
      ) {
        setVendorDropdownOpen(false);
        setVendorSearch("");
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  useEffect(() => {
    if (vendorDropdownOpen && vendorSearchRef.current) {
      vendorSearchRef.current.focus();
    }
  }, [vendorDropdownOpen]);

  const overallStatus: string = formData?.Status || "";

  const buildRibbonSteps = () => {
    const initiatorStep = {
      Role: "Initiator",
      Name: employee.EmployeeName || formData?.EmployeeName || "",
      Status: "Approved",
    };
    const approverSteps = approvalMatrix.filter((a) => a.Role !== "Initiator");
    const steps = [initiatorStep, ...approverSteps];

    if (overallStatus === "Paid") {
      return steps.map((s) => ({ ...s, _color: "approved" }));
    }
    if (overallStatus === "Reject") {
      const rejectIndex = steps.findIndex((s) => s.Status === "Reject" || s.Status === "Rejected");
      return steps.map((s, idx) => {
        if (rejectIndex === -1) return { ...s, _color: "" };
        if (idx === rejectIndex) return { ...s, _color: "rejected" };
        if (idx < rejectIndex) return { ...s, _color: "approved" };
        return { ...s, _color: "upcoming" };
      });
    }
    if (overallStatus === "Send Back" || overallStatus === "Draft") {
      return steps.map((s) =>
        s.Role === "Initiator" ? { ...s, _color: "active" } : { ...s, _color: "upcoming" },
      );
    }
    return steps.map((s) => {
      if (s.Status === "Approved") return { ...s, _color: "approved" };
      if (s.Status === "In Progress") return { ...s, _color: "active" };
      return { ...s, _color: "upcoming" };
    });
  };

  const getStepClass = (color: string) => {
    switch (color) {
      case "approved": return "approved";
      case "active": return "active";
      case "upcoming": return "upcoming";
      case "rejected": return "rejected";
      default: return "";
    }
  };
if (pageLoading) {
  return (
    <div className="MainUplodForm">
      <div className="skeletonWrapper">

        <div className="skeletonHeader"></div>

        <div className="skeletonRibbon"></div>

        {[1, 2, 3, 4, 5, 6].map((item) => (
          <div className="skeletonSection" key={item}>
            <div className="skeletonTitle"></div>

            <div className="row">
              <div className="col-md-4">
                <div className="skeletonInput"></div>
              </div>

              <div className="col-md-4">
                <div className="skeletonInput"></div>
              </div>

              <div className="col-md-4">
                <div className="skeletonInput"></div>
              </div>
            </div>
          </div>
        ))}

        <div className="skeletonTable"></div>

        <div className="skeletonButtonGroup">
          <div className="skeletonButton"></div>
          <div className="skeletonButton"></div>
          <div className="skeletonButton"></div>
        </div>

      </div>
    </div>
  );
}
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
              {buildRibbonSteps().map((a, index) => (
                <li key={index} className={`approval-step ${getStepClass(a._color)}`}>
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
                    <label className="font">Vendor Name <span className="required">*</span></label>
                    <div ref={vendorDropdownRef} style={{ position: "relative" }}>
                      <div
                        className="formtext-control"
                        onClick={() => setVendorDropdownOpen((prev) => !prev)}
                        style={{
                          cursor: "pointer",
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                          userSelect: "none",
                          minHeight: "38px",
                          padding: "6px 10px",
                          border: "1px solid #ced4da",
                          borderRadius: "4px",
                          backgroundColor: "#fff",
                        }}
                      >
                        <span style={{ color: selectedVendorName ? "inherit" : "#999" }}>
                          {selectedVendorName || "Select Vendor"}
                        </span>
                        <span style={{ fontSize: "10px", marginLeft: "8px" }}>
                          {vendorDropdownOpen ? "▲" : "▼"}
                        </span>
                      </div>

                      {vendorDropdownOpen && (
                        <div
                          style={{
                            position: "absolute",
                            top: "100%",
                            left: 0,
                            right: 0,
                            zIndex: 1000,
                            backgroundColor: "#fff",
                            border: "1px solid #ccc",
                            borderRadius: "4px",
                            boxShadow: "0 4px 8px rgba(0,0,0,0.15)",
                          }}
                        >
                          <div style={{ padding: "6px" }}>
                            <input
                              ref={vendorSearchRef}
                              type="text"
                              value={vendorSearch}
                              onChange={(e) => setVendorSearch(e.target.value)}
                              placeholder="Search vendor..."
                              className="form-control"
                              style={{ fontSize: "13px" }}
                              onClick={(e) => e.stopPropagation()}
                            />
                          </div>
                          <ul
                            style={{
                              listStyle: "none",
                              margin: 0,
                              padding: 0,
                              maxHeight: "200px",
                              overflowY: "auto",
                            }}
                          >
                            <li
                              onClick={() => {
                                setSelectedVendorId(null);
                                setSelectedVendorName("");
                                setSelectedVendorCode("");
                                setPreviousAdvances([]);
                                setVendorDropdownOpen(false);
                                setVendorSearch("");
                              }}
                              style={{
                                padding: "8px 12px",
                                cursor: "pointer",
                                color: "#999",
                                borderBottom: "1px solid #f0f0f0",
                              }}
                              onMouseEnter={(e) =>
                                ((e.currentTarget as HTMLLIElement).style.backgroundColor = "#f5f5f5")
                              }
                              onMouseLeave={(e) =>
                                ((e.currentTarget as HTMLLIElement).style.backgroundColor = "transparent")
                              }
                            >
                              Select Vendor
                            </li>
                            {filteredVendors.length === 0 ? (
                              <li style={{ padding: "8px 12px", color: "#999", fontSize: "13px" }}>
                                No vendors found
                              </li>
                            ) : (
                              filteredVendors.map((v) => (
                                <li
                                  key={v.Id}
                                  onClick={() => {
                                    setSelectedVendorId(v.Id);
                                    setSelectedVendorName(v.VendorName);
                                    setSelectedVendorCode(v.VendorCode);
                                    void getPreviousAdvances(v.Id);
                                    setVendorDropdownOpen(false);
                                    setVendorSearch("");
                                  }}
                                  style={{
                                    padding: "8px 12px",
                                    cursor: "pointer",
                                    fontSize: "13px",
                                    backgroundColor: selectedVendorId === v.Id ? "#e8f0fe" : "transparent",
                                    borderBottom: "1px solid #f0f0f0",
                                  }}
                                  onMouseEnter={(e) => {
                                    if (selectedVendorId !== v.Id)
                                      (e.currentTarget as HTMLLIElement).style.backgroundColor = "#f5f5f5";
                                  }}
                                  onMouseLeave={(e) => {
                                    if (selectedVendorId !== v.Id)
                                      (e.currentTarget as HTMLLIElement).style.backgroundColor = "transparent";
                                  }}
                                >
                                  {v.VendorName}
                                </li>
                              ))
                            )}
                          </ul>
                        </div>
                      )}
                    </div>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Code</label>
                    <input value={selectedVendorCode} className="form-control readonly" readOnly />
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
                    <label className="font">PO Basic Amount <span className="required">*</span></label>
                    <input value={poBasicAmount} onChange={(e) => handleNumberChange(e.target.value, setPoBasicAmount)} className="form-control" />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO GST Amount <span className="required">*</span></label>
                    <input value={poGstAmount} onChange={(e) => handleNumberChange(e.target.value, setPoGstAmount)} className="form-control" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Other Amount <span className="required">*</span></label>
                    <input value={poOtherAmount} onChange={(e) => handleNumberChange(e.target.value, setPoOtherAmount)} className="form-control" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Total PO Amount</label>
                    <input value={totalPoAmount ? totalPoAmount.toFixed(2) : ""} className="form-control readonly" readOnly />
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
                    <label className="font">MRN Basic Amount <span className="required">*</span></label>
                    <input value={mrnBasicAmount} onChange={(e) => handleNumberChange(e.target.value, setMrnBasicAmount)} className="form-control" />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">MRN GST Amount <span className="required">*</span></label>
                    <input value={mrnGstAmount} onChange={(e) => handleNumberChange(e.target.value, setMrnGstAmount)} className="form-control" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Other Amount <span className="required">*</span></label>
                    <input value={mrnOtherAmount} onChange={(e) => handleNumberChange(e.target.value, setMrnOtherAmount)} className="form-control" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRNAmount including GST</label>
                    <input value={totalMrnAmount ? totalMrnAmount.toFixed(2) : ""} className="form-control readonly" readOnly />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Requested Amount for Payment <span className="required">*</span></label>
                    <input value={requestedAmount} onChange={(e) => handleNumberChange(e.target.value, setRequestedAmount)} className="form-control" />
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}><label>Previous Payment Details</label></div>
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
                  {finalPayment === "No" && (
                    <div className="col-md-4">
                      <label className="font">Installation Request Number <span className="required">*</span></label>
                      <input
                        value={installationRequestNumber}
                        className="form-control"
                        onChange={(e) => setInstallationRequestNumber(e.target.value)}
                      />
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
                    <input
                      type="file"
                      multiple
                      className="form-control"
                      onChange={(e) => { if (e.target.files) setSelectedFiles(Array.from(e.target.files)); }}
                    />
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
                            .filter(
                              (h: any) =>
                                h.ActionTaken &&
                                h.ActionTaken !== "Edited" &&
                                h.ActionTaken !== "Draft Saved",
                            )
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
                <button type="button" onClick={onClose} className="reset-btn">Exit</button>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default EditAdvanceForm;