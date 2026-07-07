import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState, useRef } from "react";
import "bootstrap/dist/css/bootstrap.min.css";
import logo from "../assets/sona-comstarlogo.png";
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

const NewAdvanceform = ({ context, onClose }: any) => {
  const sp = spfi().using(SPFx(context));
  const today = new Date();
  const localDate: string = new Date(
    today.getTime() - today.getTimezoneOffset() * 60000,
  )
    .toISOString()
    .split("T")[0];

  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const tenantUrl = context.pageContext.site.absoluteUrl.split("/sites/")[0];
  const vendorSp = spfi(`${tenantUrl}/sites/RLY_AccountsPayable_UAT`).using(SPFx(context));
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isDraftSaving, setIsDraftSaving] = useState(false);
  const [employee, setEmployee] = React.useState<any>({});
  const [attachments, setAttachments] = useState<File[]>([]);
  const [previousAdvances, setPreviousAdvances] = useState<IPreviousAdvance[]>([]);
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [finalPayment, setFinalPayment] = useState("");
  const [installationDetails, setInstallationDetails] = useState("");
  const [poBasicAmount, setPoBasicAmount] = useState("");
  const [poGstAmount, setPoGstAmount] = useState("");
  const [poOtherAmount, setPoOtherAmount] = useState("");
  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [mrnNumber, setMrnNumber] = useState("");
  const [mrnDate, setMrnDate] = useState("");
  const [mrnBasicAmount, setMrnBasicAmount] = useState("");
  const [mrnGstAmount, setMrnGstAmount] = useState("");
  const [mrnOtherAmount, setMrnOtherAmount] = useState("");
  const [requestedAmount, setRequestedAmount] = useState("");
  const [requesterRemarks, setRequesterRemarks] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [installationRequestNumber, setInstallationRequestNumber] = useState("");

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

  const employeeRef = useRef<any>({});

  const totalPoAmount =
    (Number(poBasicAmount) || 0) +
    (Number(poGstAmount) || 0) +
    (Number(poOtherAmount) || 0);

  const totalMrnAmount =
    (Number(mrnBasicAmount) || 0) +
    (Number(mrnGstAmount) || 0) +
    (Number(mrnOtherAmount) || 0);

  const toNumber = (val: string): number | null =>
    val && val.trim() !== "" ? parseFloat(val) : null;

  const toTotalOrNull = (a: string, b: string, c: string): number | null => {
    if (
      (!a || a.trim() === "") &&
      (!b || b.trim() === "") &&
      (!c || c.trim() === "")
    )
      return null;
    return (Number(a) || 0) + (Number(b) || 0) + (Number(c) || 0);
  };

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

      // const FLOW_URL =
      //   "https://defaultcb1edbfe8080457d9cae51528f3643.3f.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/e2bb522aa41443179a72b701b9613471/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=q8b8ADCtK2eKr2f6p3MX7gxmJymPeJbm0mq2M69Rk8E";

      // const fetchPage = async (pageNumber: number) => {
      //   const response = await fetch(FLOW_URL, {
      //     method: "POST",
      //     headers: {
      //       "Content-Type": "application/json",
      //     },
      //     body: JSON.stringify({
      //       PageSize: 500,
      //       PageNumber: pageNumber,
      //     }),
      //   });

      //   if (!response.ok) {
      //     throw new Error("Failed to fetch employee data");
      //   }

      //   return response.json();
      // };
        const fetchPage = async (pageNumber: number) => {
       
        const username = "0le867nyvalvfo249e6sj4ri";
        const password = "2mpvr7r19amf7o01hr0qncr861hmtsb7o9ap51hwar72405atj3y73mndkmokg5i";

        const auth = btoa(`${username}:${password}`);

        const responsesevices = await fetch(
          "https://mservices.zinghr.com/etl/api/v2/Auth/GenerateJWTToken?apiPermission=GEMD",
          {
            method: "GET",
            headers: {
              "Authorization": `Basic ${auth}`,
              "Accept": "application/json",
              "Content-Type": "application/json",
            },
          }
        );

        const ServicedataToken = await responsesevices.json();
        console.log(ServicedataToken.data);

         const response = await fetch("https://mservices.zinghr.com/etl/api/v2/Employee/GetEmployeeDetails", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${ServicedataToken.data}`,
            "ClientSecret": "2mpvr7r19amf7o01hr0qncr861hmtsb7o9ap51hwar72405atj3y73mndkmokg5i"
          },
          body: JSON.stringify({
            PageSize: 500,
            PageNumber: pageNumber,
          }),
        });
        // const data = await response.json();
        // console.log("Employee data:", data);
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

      employeeRef.current = {
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
      buildApprovalPreview({
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

  const getVendors = async () => {
    try {
      const data = await vendorSp.web.lists
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
    return (month >= 4 ? year + 1 : year).toString().slice(-2);
  };

  const generatePaymentId = async (): Promise<string> => {
    try {
      const fy = getFinancialYear();
      const incrementalData = await sp.web.lists
        .getByTitle("CapexPaymentIncrementalID")
        .items.select("Id", "IncrementalID")
        .orderBy("Id", false)
        .top(1)();
      const nextNumber =
        incrementalData.length > 0 && incrementalData[0].IncrementalID
          ? Number(incrementalData[0].IncrementalID) + 1
          : 1;

      await sp.web.lists.getByTitle("CapexPaymentIncrementalID").items.getById(incrementalData[0].Id).update({
        //Title: `CPX-${nextNumber}`,
        IncrementalID: nextNumber.toString(),
      });
      return `CPX-PYMT/${fy}/${nextNumber.toString().padStart(5, "0")}`;
    } catch (error) {
      console.error("Generate Payment ID Error:", error);
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
        Status: "Pending"
      });
    if (emp.HOD?.Title)
      flow.push({
        Name: emp.HOD.Title,
        Role: "HOD",
        Status: "Pending"
      });
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

  const buildApprovalFlow = async (emp: any) => {
    const flow: any[] = [];
    if (emp.ReportingManager?.Id)
      flow.push({
        Id: emp.ReportingManager.Id,
        Name: emp.ReportingManager.Title,
        Role: "RM",
        Level: 1,
        Status: "Pending",
      });
    if (emp.HOD?.Id)
      flow.push({
        Id: emp.HOD.Id,
        Name: emp.HOD.Title,
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
    if (!poBasicAmount || poBasicAmount.trim() === "")
      errors.push("Please enter PO Basic Amount");
    if (!poGstAmount || poGstAmount.trim() === "")
      errors.push("Please enter PO GST Amount");
    if (!poOtherAmount || poOtherAmount.trim() === "")
      errors.push("Please enter PO Other Amount");
    if (!mrnNumber || mrnNumber.trim() === "")
      errors.push("Please enter MRN Number");
    if (!mrnDate || mrnDate === "Invalid Date")
      errors.push("Please select MRN Date");
    if (mrnDate > localDate) errors.push("MRN date cannot be a future date");
    if (!mrnBasicAmount || mrnBasicAmount.trim() === "")
      errors.push("Please enter MRN Basic Amount");
    if (!mrnGstAmount || mrnGstAmount.trim() === "")
      errors.push("Please enter MRN GST Amount");
    if (!mrnOtherAmount || mrnOtherAmount.trim() === "")
      errors.push("Please enter MRN Other Amount");
    if (!requestedAmount || requestedAmount.trim() === "")
      errors.push("Please enter Requested Amount");
    if (finalPayment === "Yes" && !installationDetails)
      errors.push("Please enter Installation Details");
    if (
      finalPayment === "No" &&
      (!installationRequestNumber || installationRequestNumber.trim() === "")
    )
      errors.push("Please enter Installation Request Number");
    if (!attachments || attachments.length === 0)
      errors.push("Please upload attachment");
    if (requestedAmount && totalMrnAmount && Number(requestedAmount) > totalMrnAmount)
      errors.push("Requested Amount should not be greater than MRN Amount (GST)");
    return errors;
  };

  const buildPayload = (
    emp: any,
    capexId: string,
    flow: any[],
    status: string,
    workflowHistory: any[],
    isDraft: boolean,
  ) => {
    const currentApprover = flow.length > 0 ? flow[0].Id : null;

    const poBasic = isDraft ? toNumber(poBasicAmount) : parseFloat(poBasicAmount);
    const poGst = isDraft ? toNumber(poGstAmount) : parseFloat(poGstAmount);
    const poOther = isDraft ? toNumber(poOtherAmount) : parseFloat(poOtherAmount);
    const mrnBasic = isDraft ? toNumber(mrnBasicAmount) : parseFloat(mrnBasicAmount);
    const mrnGst = isDraft ? toNumber(mrnGstAmount) : parseFloat(mrnGstAmount);
    const mrnOther = isDraft ? toNumber(mrnOtherAmount) : parseFloat(mrnOtherAmount);

    const poTotal = isDraft
      ? toTotalOrNull(poBasicAmount, poGstAmount, poOtherAmount)
      : totalPoAmount || null;

    const mrnTotal = isDraft
      ? toTotalOrNull(mrnBasicAmount, mrnGstAmount, mrnOtherAmount)
      : totalMrnAmount || null;

    const reqAmount = isDraft
      ? (requestedAmount && requestedAmount.trim() !== "" ? requestedAmount : null)
      : requestedAmount || null;

    return {
      Title: capexId,
      CapexId: capexId,
      EmployeeCode: emp.EmployeeCode || null,
      EmployeeName: emp.EmployeeName || null,
      Division: emp.Division || null,
      Location: emp.Location || null,
      Email: emp.EmployeeEmail || null,
      RM: emp.ReportingManager?.Title || null,
      HOD: emp.HOD?.Title || null,
      ContactNo: emp.ContactNo || null,
      EmployeeStatus: emp.EmployeeStatus || null,
      VendorCode: selectedVendorCode || null,
      VendorName: selectedVendorName || null,
      PONumber: poNumber || null,
      PODate: poDate ? new Date(poDate) : null,
      POPaymentTerms: poTerms || null,
      POBasicAmount: poBasic,
      POGSTAmount: poGst,
      POOtherAmount: poOther,
      POAmount: poTotal !== null ? poTotal.toString() : null,
      MRNNumber: mrnNumber || null,
      MRNDtae: mrnDate ? new Date(mrnDate) : null,
      MRNBasicAmount: mrnBasic,
      MRNGSTAmount: mrnGst,
      MRNOtherAmount: mrnOther,
      MRNAmountwithGST: mrnTotal !== null ? mrnTotal.toString() : null,
      RequestedAmountforPayment: reqAmount,
      FinalPaymentAgainstPO: finalPayment === "Yes" ? true : finalPayment === "No" ? false : null,
      InstallationDetails: installationDetails || null,
      InstallationRequestNumber: finalPayment === "No" ? installationRequestNumber || null : null,
      RequesterRemarks: requesterRemarks || null,
      StatusFlow: status,
      Status: status,
      ApprovalMatrix: JSON.stringify(flow),
      CurrentApproverId: currentApprover,
      WorkflowHistory: JSON.stringify(workflowHistory),
    };
  };

  const handleSubmit = async () => {
    if (isSubmitting) return;
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
      const emp = employeeRef.current;
      const capexId = await generatePaymentId();
      const flow = await buildApprovalFlow(emp);
      const WorkflowHistory = [
        {
          CurrentApprover: emp.EmployeeName,
          ActionTaken: "Submitted",
          Comment: requesterRemarks,
          Date: new Date().toISOString(),
        },
      ];
      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.add(buildPayload(emp, capexId, flow, "Pending for Approval", WorkflowHistory, false));
      await uploadAttachments(capexId);
      await Swal.fire({
        icon: "success",
        title: "Success",
        text: "Submitted successfully.",
        confirmButtonText: "OK",
      });
      onClose();
    } catch (error: any) {
      console.error("FULL ERROR:", error);
      await Swal.fire({
        icon: "error",
        title: "Submission Failed",
        text: error?.data?.responseBody || error?.message || "Error while saving.",
        confirmButtonText: "OK",
      });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handledraft = async () => {
    if (isDraftSaving) return;
    setIsDraftSaving(true);
    try {
      const emp = employeeRef.current;
      if (!emp || !emp.EmployeeName) {
        await Swal.fire({
          icon: "error",
          title: "Not Ready",
          text: "Employee information is still loading. Please wait a moment and try again.",
          confirmButtonText: "OK",
        });
        return;
      }
      const capexId = await generatePaymentId();
      const flow = await buildApprovalFlow(emp);
      flow.forEach((f: any) => (f.Status = "Pending"));
      const WorkflowHistory = [
        {
          CurrentApprover: emp.EmployeeName,
          Comment: requesterRemarks || "",
          Date: new Date().toISOString(),
        },
      ];
      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.add(buildPayload(emp, capexId, flow, "Draft", WorkflowHistory, true));
      await uploadAttachments(capexId);
      await Swal.fire({
        icon: "success",
        title: "Success",
        text: "Draft saved successfully.",
        confirmButtonText: "OK",
      });
      onClose();
    } catch (error: any) {
      console.error("Draft save ERROR:", error);
      await Swal.fire({
        icon: "error",
        title: "Save Failed",
        text: error?.data?.responseBody || error?.message || "Error while saving draft.",
        confirmButtonText: "OK",
      });
    } finally {
      setIsDraftSaving(false);
    }
  };

  React.useEffect(() => {
    if (!context) return;

    void loadPage();

  }, [context]);

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

  if (pageLoading) {

    return (

      <div className="skeletonWrapper">
        <div className="skeletonLine" style={{ width: "40%" }} />
        <div className="skeletonLine" style={{ width: "80%" }} />

        <div className="skeletonCard"></div>

        <div className="keletonLine" style={{ width: "60%" }} />
        <div className="skeletonCard"></div>

        <div className="skeletonLine" style={{ width: "50%" }} />
        <div className="skeletonCard"></div>
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
              <h1>Capex Payment Request</h1>
            </div>

            {approvalMatrix.length === 0 ? (
              <p>Loading...</p>
            ) : (
              <div className="displayWF">
                <ul className="approval-flow">
                  <li className="approval-step active">
                    Initiator - {employee.EmployeeName}
                  </li>
                  {approvalMatrix.map((a, index) => (
                    <li key={index} className="approval-step upcoming">
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
                    <label className="font">Employee Email</label> : &nbsp;&nbsp;
                    <label className="fonttext">{employee.EmployeeEmail}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Contact No</label> : &nbsp;&nbsp;
                    <label className="fonttext">{employee.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Status</label> : &nbsp;&nbsp;
                    <label className="fonttext">{employee.EmployeeStatus}</label>
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
                    <label className="fonttext">{employee.ReportingManager?.Title}</label>
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
                      Vendor Name&nbsp;<span className="required">*</span>
                    </label>
                    <div
                      ref={vendorDropdownRef}
                      style={{ position: "relative" }}
                    >
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
                                setPoNumber("");
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
                              <li
                                style={{
                                  padding: "8px 12px",
                                  color: "#999",
                                  fontSize: "13px",
                                }}
                              >
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
                                    setPoNumber("");
                                    void getPreviousAdvances(v.Id);
                                    setVendorDropdownOpen(false);
                                    setVendorSearch("");
                                  }}
                                  style={{
                                    padding: "8px 12px",
                                    cursor: "pointer",
                                    fontSize: "13px",
                                    backgroundColor:
                                      selectedVendorId === v.Id ? "#e8f0fe" : "transparent",
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
                    <input
                      value={selectedVendorCode}
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
                      PO Payment Terms&nbsp;<span className="required">*</span>
                    </label>
                    <input
                      value={poTerms}
                      className="form-control"
                      onChange={(e) => setPoTerms(e.target.value)}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">
                      PO Basic Amount&nbsp;<span className="required">*</span>
                    </label>
                    <input
                      value={poBasicAmount}
                      className="form-control"
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setPoBasicAmount)
                      }
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">
                      PO GST Amount&nbsp;<span className="required">*</span>
                    </label>
                    <input
                      value={poGstAmount}
                      className="form-control"
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setPoGstAmount)
                      }
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">
                      PO Other Amount&nbsp;<span className="required">*</span>
                    </label>
                    <input
                      value={poOtherAmount}
                      className="form-control"
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setPoOtherAmount)
                      }
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Total PO Amount</label>
                    <input
                      value={totalPoAmount ? totalPoAmount.toFixed(2) : ""}
                      className="form-control readonly"
                      readOnly
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
                      MRN Basic Amount&nbsp;<span className="required">*</span>
                    </label>
                    <input
                      value={mrnBasicAmount}
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setMrnBasicAmount)
                      }
                      className="form-control"
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">
                      MRN GST Amount&nbsp;<span className="required">*</span>
                    </label>
                    <input
                      value={mrnGstAmount}
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setMrnGstAmount)
                      }
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">
                      MRN Other Amount&nbsp;<span className="required">*</span>
                    </label>
                    <input
                      value={mrnOtherAmount}
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setMrnOtherAmount)
                      }
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Total MRN Amount</label>
                    <input
                      value={totalMrnAmount ? totalMrnAmount.toFixed(2) : ""}
                      className="form-control readonly"
                      readOnly
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
                            `${context.pageContext.web.absoluteUrl}/SitePages/InstallationComm.aspx`,
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
                      <label className="font">
                        Installation Request Number&nbsp;
                        <span className="required">*</span>
                      </label>
                      <input
                        value={installationRequestNumber}
                        className="form-control"
                        onChange={(e) =>
                          setInstallationRequestNumber(e.target.value)
                        }
                      />
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
                              previousAdvances.map((item: any, index: number) => {
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
                                        ? new Date(item.Created).toLocaleDateString()
                                        : ""}
                                    </td>
                                    <td>
                                      {item.VoucherDate
                                        ? new Date(item.VoucherDate).toLocaleDateString()
                                        : ""}
                                    </td>
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
                    <button
                      type="button"
                      onClick={onClose}
                      className="reset-btn"
                    >
                      Exit
                    </button>
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