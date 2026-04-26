import * as React from 'react';
import styles from './ApplyLeave.module.scss';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/site-users/web";
import EditLeave from './EditLeave';

interface IFormState {
  leaveTypeId: number;
  leaveDurationType: 'Full Day' | 'Half Day';
  startDate: string;
  endDate: string;
  totalDays: number;
  halfDayType: 'First Half' | 'Second Half';
  reason: string;
  newAttachments: File[];  
  otherLeaveType: string;
}

interface IFormErrors {
  leaveType?: string;
  startDate?: string;
  endDate?: string;
  reason?: string;
  overlap?: string;
  attachment?: string;
  submit?: string;
  weekend?: string;
  manager?: string;
  probation?: string;
  otherLeaveType?: string;
  insufficientQuota?: string;
}

interface IEmployeeData {
  id: number;
  name: string;
  department: string;
  departmentId: number;
  country: string;
  countryId: number;
  countryCode: string;
  countryCodeId: number;
  managerId: number;
  managerName: string;
  managerEmail: string;
  joinDate: string;
  employmentType: string;
  userImageUrl: string;
}

interface ILeaveType {
  Id: number;
  Title: string;
}

interface ILeaveQuota {
  Id: number;
  CountryId: number;
  CountryCodeId: number;
  Leaves: number;
  AnnualLeaves: number;
  CasualLeaves: number;
  SickLeaves: number;
  OtherLeaves: number;
}

interface IEmployeeLeaveBalance {
  Id: number;
  LeavesBalance: number;
  Used: number;
  Remaining: number;
  Year: string;
}

interface ILeaveCycle {
  startDate: Date;
  endDate: Date;
  cycleNumber: number;
}

interface ILeaveRequest {
  Id: number;
  StartDate: string;
  EndDate: string;
  Status: string;
  TotalDays: number;
  LeaveTypeId: number;
  EmployeeYearCycle: number;
}

interface IApplyLeaveState {
  formState: IFormState;
  errors: IFormErrors;
  isSubmitting: boolean;
  showProbationPopup: boolean;
  showConfirmationPopup: boolean;
  showSuccessPopup: boolean;
  weekendError: string | null;
  showHalfDayType: boolean;
  currentUser: any;
  leaveQuota: ILeaveQuota | null;
  employeeLeaveBalance: IEmployeeLeaveBalance | null;
  currentCycle: ILeaveCycle | null;
  usedLeavesInCycle: {
    AnnualLeaves: number;
    CasualLeaves: number;
    SickLeaves: number;
    OtherLeaves: number;
  };
  employee: IEmployeeData | null;
  loading: boolean;
  existingRequests: ILeaveRequest[] | null;
  showEditLeave: boolean;
  initialLoadComplete: boolean;
}

const STATIC_LEAVE_TYPES: ILeaveType[] = [
  { Id: 1, Title: 'Annual Leave' },
  { Id: 2, Title: 'Casual Leave' },
  { Id: 3, Title: 'Sick Leave' },
  { Id: 4, Title: 'Other Leave' }
];

const FORM_INITIAL_STATE: IFormState = {
  leaveTypeId: 0,
  leaveDurationType: 'Full Day',
  startDate: '',
  endDate: '',
  totalDays: 0,
  halfDayType: 'First Half',
  reason: '',
  newAttachments: [],
  otherLeaveType: '',
};

const FILE_CONFIG = {
  MAX_SIZE: 5 * 1024 * 1024,
  ACCEPTED_TYPES: ['application/pdf', 'image/jpeg', 'image/png', 'image/jpg'] as const,
};

type AcceptedFileType = typeof FILE_CONFIG.ACCEPTED_TYPES[number];

const formatDate = (date: Date): string => date.toISOString().split('T')[0];
const isWeekend = (date: Date): boolean => date.getDay() === 5 || date.getDay() === 6;

const calculateTotalDays = (startDate: Date, endDate: Date, isHalfDay: boolean): number => {
  if (isHalfDay) return 0.5;
  let days = 0;
  const current = new Date(startDate);
  const end = new Date(endDate);
  while (current <= end) {
    days++;
    current.setDate(current.getDate() + 1);
  }
  return days;
};

const validateWeekendDates = (startDate: Date, endDate: Date): string | null => {
  if (isWeekend(startDate)) return 'Start date cannot be a weekend (Friday or Saturday)';
  if (isWeekend(endDate)) return 'End date cannot be a weekend (Friday or Saturday)';
  return null;
};

const isDateOverlap = (start1: Date, end1: Date, start2: Date, end2: Date): boolean =>
  start1 <= end2 && end1 >= start2;

const isEmployeeEligibleForLeaveQuota = (employmentType: string): boolean => {
  const eligibleTypes = ['Permanent', 'Contract', 'Part-time'];
  return eligibleTypes.includes(employmentType);
};

const calculateCurrentLeaveCycle = (joinDate: string): ILeaveCycle => {
  const join = new Date(joinDate);
  const today = new Date();

  let cycleStart = new Date(join);
  let cycleEnd = new Date(join);
  cycleEnd.setFullYear(cycleEnd.getFullYear() + 1);
  cycleEnd.setDate(cycleEnd.getDate() - 1);

  let cycleNumber = 1;

  while (today > cycleEnd) {
    cycleStart = new Date(cycleEnd);
    cycleStart.setDate(cycleStart.getDate() + 1);
    cycleEnd = new Date(cycleStart);
    cycleEnd.setFullYear(cycleEnd.getFullYear() + 1);
    cycleEnd.setDate(cycleEnd.getDate() - 1);
    cycleNumber++;
  }

  return { startDate: cycleStart, endDate: cycleEnd, cycleNumber };
};

// Simple in-memory cache
const simpleCache = new Map<string, any>();

const getCached = (key: string): any => {
  const cached = simpleCache.get(key);
  if (cached && Date.now() - cached.timestamp < 300000) {
    return cached.data;
  }
  return null;
};

const setCache = (key: string, data: any): void => {
  simpleCache.set(key, { data, timestamp: Date.now() });
};

export default class ApplyLeave extends React.Component<{ onLeaveSubmitted?: () => void }, IApplyLeaveState> {
  private isMounted = false;

  constructor(props: { onLeaveSubmitted?: () => void }) {
    super(props);
    this.state = {
      formState: FORM_INITIAL_STATE,
      errors: {},
      isSubmitting: false,
      showProbationPopup: false,
      showConfirmationPopup: false,
      showSuccessPopup: false,
      weekendError: null,
      showHalfDayType: false,
      currentUser: null,
      leaveQuota: null,
      employeeLeaveBalance: null,
      currentCycle: null,
      usedLeavesInCycle: {
        AnnualLeaves: 0,
        CasualLeaves: 0,
        SickLeaves: 0,
        OtherLeaves: 0
      },
      employee: null,
      loading: true,
      existingRequests: null,
      showEditLeave: false,
      initialLoadComplete: false
    };
  }

  checkUrlForRequestId = (): boolean => {
    const urlParams = new URLSearchParams(window.location.search);
    const requestIdParam = urlParams.get('RequestID');

    if (requestIdParam && requestIdParam.trim() !== '') {
      const requestId = parseInt(requestIdParam, 10);
      if (!isNaN(requestId) && requestId > 0) {
        return true;
      }
    }
    return false;
  };

  async componentDidMount() {
    this.isMounted = true;

    const hasValidRequestId = this.checkUrlForRequestId();

    if (hasValidRequestId) {
      if (this.isMounted) {
        this.setState({ showEditLeave: true, loading: false });
      }
    } else {
      if (typeof window !== 'undefined' && 'requestIdleCallback' in window) {
        requestIdleCallback(() => this.fetchAllDataOptimized());
      } else {
        setTimeout(() => this.fetchAllDataOptimized(), 50);
      }
    }
  }

  componentWillUnmount() {
    this.isMounted = false;
  }

  fetchAllDataOptimized = async () => {
    if (!this.isMounted) return;

    try {
      const cachedUser = getCached('currentUser');
      let user;
      
      if (cachedUser) {
        user = cachedUser;
      } else {
        user = await sp.web.currentUser();
        setCache('currentUser', user);
      }
      
      if (!this.isMounted) return;

      this.setState({ currentUser: user });

      const employeeData = await this.fetchEmployeeDataFast(user.Email);
      
      if (!this.isMounted) return;

      if (!employeeData) {
        throw new Error('No employee record found');
      }

      const deptData = await this.fetchDepartmentDataFast();
      const employee = this.processEmployeeDataBasic(employeeData, user, deptData);
      const cycle = calculateCurrentLeaveCycle(employee.joinDate);

      if (this.isMounted) {
        this.setState({
          employee,
          currentCycle: cycle,
          loading: false,
          initialLoadComplete: true
        });
      }

      this.fetchRemainingDataInBackground(employee, cycle);
      setTimeout(() => this.checkProbationStatus(), 100);

    } catch (err) {
      if (this.isMounted) {
        this.setState({ 
          loading: false, 
          errors: { submit: 'Failed to load data. Please refresh the page.' } 
        });
      }
    }
  };

  fetchRemainingDataInBackground = async (employee: IEmployeeData, cycle: ILeaveCycle) => {
    try {
      const [quota, balance, requests] = await Promise.all([
        this.fetchLeaveQuotaFast(employee.countryId, employee.countryCodeId),
        this.fetchLeaveBalanceFast(employee.id),
        this.fetchLeaveRequestsFast(employee.id, cycle.cycleNumber)
      ]);

      if (!this.isMounted) return;

      const usedLeaves = this.calculateUsedLeavesFast(requests);

      this.setState({
        leaveQuota: quota,
        employeeLeaveBalance: balance,
        existingRequests: requests,
        usedLeavesInCycle: usedLeaves
      });

    } catch (err) {
      console.warn('Background data fetch failed:', err);
    }
  };

  fetchEmployeeDataFast = async (email: string): Promise<any> => {
    const cacheKey = `employee_${email}`;
    const cached = getCached(cacheKey);
    if (cached) return cached;

    try {
      const items = await sp.web.lists
        .getByTitle('Employee Information List')
        .items
        .select(
          'Id',
          'EmployeeName/Id',
          'EmployeeName/Title',
          'EmployeeName/EMail',
          'Department/Id',
          'Department/Title',
          'Country/Id',
          'Country/Title',
          'CountryCode/Id',
          'CountryCode/Title',
          'JoiningDate',
          'EmploymentType'
        )
        .expand('EmployeeName', 'Department', 'Country', 'CountryCode')
        .filter(`EmployeeName/EMail eq '${email.replace(/'/g, "''")}'`)
        .top(1)
        .get();

      const result = items && items.length > 0 ? items[0] : null;
      if (result) setCache(cacheKey, result);
      return result;
    } catch {
      return null;
    }
  };

  fetchDepartmentDataFast = async (): Promise<Map<number, any>> => {
    const cacheKey = 'all_departments';
    const cached = getCached(cacheKey);
    if (cached) return cached;

    try {
      const items = await sp.web.lists
        .getByTitle('Department')
        .items
        .select('Id', 'Title', 'DepartmentManager/Id', 'DepartmentManager/Title', 'DepartmentManager/EMail')
        .expand('DepartmentManager')
        .get();
      
      const deptMap = new Map();
      items.forEach(item => {
        deptMap.set(item.Id, item);
      });
      
      setCache(cacheKey, deptMap);
      return deptMap;
    } catch {
      return new Map();
    }
  };

  processEmployeeDataBasic = (empData: any, user: any, deptMap: Map<number, any>): IEmployeeData => {
    const departmentId = empData.Department?.Id;
    const deptData = departmentId ? deptMap.get(departmentId) : null;

    let managerId = 0, managerName = '', managerEmail = '';
    if (deptData?.DepartmentManager) {
      managerId = deptData.DepartmentManager.Id;
      managerName = deptData.DepartmentManager.Title;
      managerEmail = deptData.DepartmentManager.EMail;
    }

    return {
      id: user?.Id || 0,
      name: empData.EmployeeName?.Title || user?.Title || '',
      department: empData.Department?.Title || '',
      departmentId: departmentId || 0,
      country: empData.Country?.Title || '',
      countryId: empData.Country?.Id || 0,
      countryCode: empData.CountryCode?.Title || '',
      countryCodeId: empData.CountryCode?.Id || 0,
      managerId, managerName, managerEmail,
      joinDate: empData.JoiningDate || '',
      employmentType: empData.EmploymentType || 'Permanent',
      userImageUrl: '',
    };
  };

  fetchLeaveQuotaFast = async (countryId: number, countryCodeId: number): Promise<ILeaveQuota | null> => {
    if (!countryId || !countryCodeId) return null;

    const cacheKey = `leave_quota_${countryId}_${countryCodeId}`;
    const cached = getCached(cacheKey);
    if (cached) return cached;

    try {
      const items = await sp.web.lists
        .getByTitle('Leave Quota')
        .items
        .select(
          'Id',
          'Country/Id',
          'CountryCode/Id',
          'Leaves',
          'AnnualLeaves',
          'CasualLeaves',
          'SickLeaves',
          'OtherLeaves'
        )
        .expand('Country', 'CountryCode')
        .filter(`Country/Id eq ${countryId} and CountryCode/Id eq ${countryCodeId}`)
        .top(1)
        .get();

      if (items && items.length > 0) {
        const quota = {
          Id: items[0].Id,
          CountryId: items[0].Country?.Id,
          CountryCodeId: items[0].CountryCode?.Id,
          Leaves: items[0].Leaves,
          AnnualLeaves: items[0].AnnualLeaves,
          CasualLeaves: items[0].CasualLeaves,
          SickLeaves: items[0].SickLeaves,
          OtherLeaves: items[0].OtherLeaves
        } as ILeaveQuota;
        setCache(cacheKey, quota);
        return quota;
      }
      return null;
    } catch {
      return null;
    }
  };

  fetchLeaveRequestsFast = async (employeeId: number, cycleNumber: number): Promise<ILeaveRequest[]> => {
    if (!employeeId) return [];

    const cacheKey = `leave_requests_${employeeId}_${cycleNumber}`;
    const cached = getCached(cacheKey);
    if (cached) return cached;

    try {
      const filter = `EmployeeId eq ${employeeId} and (Status eq 'Pending on Manager' or Status eq 'Pending on HR' or Status eq 'Pending on Executive' or Status eq 'Approved') and EmployeeYearCycle eq ${cycleNumber}`;

      const items = await sp.web.lists
        .getByTitle('Leave Requests')
        .items
        .select('Id', 'StartDate', 'EndDate', 'Status', 'TotalDays', 'LeaveTypeId', 'EmployeeYearCycle')
        .filter(filter)
        .orderBy('StartDate', false)
        .get();

      setCache(cacheKey, items);
      return items as ILeaveRequest[];
    } catch {
      return [];
    }
  };

  fetchLeaveBalanceFast = async (employeeId: number): Promise<IEmployeeLeaveBalance | null> => {
    if (!employeeId) return null;

    const currentYear = new Date().getFullYear().toString();
    const cacheKey = `leave_balance_${employeeId}_${currentYear}`;
    const cached = getCached(cacheKey);
    if (cached) return cached;

    try {
      const items = await sp.web.lists
        .getByTitle('Employee Leave Balance')
        .items
        .select('Id', 'LeavesBalance', 'Used', 'Remaining', 'Year', 'EmployeeId')
        .filter(`EmployeeId eq ${employeeId} and Year eq '${currentYear}'`)
        .top(1)
        .get();

      const result = items && items.length > 0 ? items[0] as IEmployeeLeaveBalance : null;
      if (result) setCache(cacheKey, result);
      return result;
    } catch {
      return null;
    }
  };

  calculateUsedLeavesFast = (requests: ILeaveRequest[] | null) => {
    const used = { AnnualLeaves: 0, CasualLeaves: 0, SickLeaves: 0, OtherLeaves: 0 };

    if (requests && requests.length > 0) {
      for (const request of requests) {
        const totalDays = request.TotalDays || 0;
        if (request.LeaveTypeId === 1) used.AnnualLeaves += totalDays;
        else if (request.LeaveTypeId === 2) used.CasualLeaves += totalDays;
        else if (request.LeaveTypeId === 3) used.SickLeaves += totalDays;
        else if (request.LeaveTypeId === 4) used.OtherLeaves += totalDays;
      }
    }

    return used;
  };

  checkProbationStatus = () => {
    const { employee } = this.state;
    if (!employee?.employmentType) return;

    const isProbation = employee.employmentType.toLowerCase() === 'probation';

    if (isProbation && this.isMounted) {
      this.setState({ showProbationPopup: true });
    }
  };

  isProbation = (): boolean => {
    const { employee } = this.state;
    if (!employee?.employmentType) return false;
    return employee.employmentType.toLowerCase() === 'probation';
  };

  getAvailableQuota = (leaveTypeId: number): number => {
    const { leaveQuota, usedLeavesInCycle, employeeLeaveBalance, currentCycle } = this.state;

    if (this.isProbation()) return 0;
    if (!isEmployeeEligibleForLeaveQuota(this.state.employee?.employmentType || '')) return 0;
    if (!leaveQuota) return 0;

    let baseQuota = 0;
    switch (leaveTypeId) {
      case 1:
        baseQuota = leaveQuota.AnnualLeaves;
        break;
      case 2:
        baseQuota = leaveQuota.CasualLeaves;
        break;
      case 3:
        baseQuota = leaveQuota.SickLeaves;
        break;
      case 4:
        baseQuota = leaveQuota.OtherLeaves;
        break;
      default:
        return 0;
    }

    let totalAvailable = baseQuota;

    if (leaveTypeId !== 4 && currentCycle && currentCycle.cycleNumber > 1 && employeeLeaveBalance) {
      totalAvailable = baseQuota + (employeeLeaveBalance.Remaining || 0);
    }

    let usedLeaves = 0;
    switch (leaveTypeId) {
      case 1:
        usedLeaves = usedLeavesInCycle.AnnualLeaves;
        break;
      case 2:
        usedLeaves = usedLeavesInCycle.CasualLeaves;
        break;
      case 3:
        usedLeaves = usedLeavesInCycle.SickLeaves;
        break;
      case 4:
        usedLeaves = usedLeavesInCycle.OtherLeaves;
        break;
    }

    return Math.max(0, totalAvailable - usedLeaves);
  };

  canApplyForLeave = (leaveTypeId: number, requestedDays: number): boolean => {
    if (this.isProbation()) return false;
    if (!isEmployeeEligibleForLeaveQuota(this.state.employee?.employmentType || '')) return false;

    const available = this.getAvailableQuota(leaveTypeId);
    return requestedDays <= available;
  };

  validateForm = (): IFormErrors => {
    const { formState, existingRequests } = this.state;
    const newErrors: IFormErrors = {};

    if (this.isProbation()) {
      newErrors.probation = 'You are on probation period. You cannot apply for any leave.';
      return newErrors;
    }

    if (!formState.leaveTypeId || formState.leaveTypeId === 0) newErrors.leaveType = 'Please select a leave type';
    if (formState.leaveTypeId === 4 && !formState.otherLeaveType.trim()) newErrors.otherLeaveType = 'Please specify other leave type';

    if (!formState.startDate) {
      newErrors.startDate = 'Please select start date';
    } else {
      const startDate = new Date(formState.startDate);
      if (isWeekend(startDate)) newErrors.startDate = 'Start date cannot be a weekend (Friday or Saturday)';
    }

    if (formState.leaveDurationType === 'Full Day' && !formState.endDate) {
      newErrors.endDate = 'Please select end date';
    } else if (formState.startDate && formState.endDate) {
      const start = new Date(formState.startDate);
      const end = new Date(formState.endDate);
      if (end < start) newErrors.endDate = 'End date cannot be before start date';
      if (isWeekend(end)) newErrors.endDate = 'End date cannot be a weekend (Friday or Saturday)';
    }

    if (!formState.reason) {
      newErrors.reason = 'Please provide a reason for leave';
    } else if (formState.reason.length < 10) {
      newErrors.reason = 'Please provide a more detailed reason (minimum 10 characters)';
    }

    if (existingRequests && existingRequests.length > 0 && formState.startDate && formState.endDate) {
      const newStart = new Date(formState.startDate);
      const newEnd = new Date(formState.endDate);
      const hasOverlap = existingRequests.some((leave: ILeaveRequest) => {
        const existingStart = new Date(leave.StartDate);
        const existingEnd = new Date(leave.EndDate);
        return isDateOverlap(newStart, newEnd, existingStart, existingEnd);
      });
      if (hasOverlap) newErrors.overlap = 'You already have a leave request for these dates';
    }

    if (formState.leaveTypeId && formState.leaveTypeId !== 4 && formState.totalDays > 0) {
      const canApply = this.canApplyForLeave(formState.leaveTypeId, formState.totalDays);
      const selectedLeaveType = STATIC_LEAVE_TYPES.find(type => type.Id === formState.leaveTypeId);
      const availableQuota = this.getAvailableQuota(formState.leaveTypeId);

      if (!canApply) {
        newErrors.insufficientQuota = `Insufficient ${selectedLeaveType?.Title} quota. Available: ${availableQuota} days, Requested: ${formState.totalDays} days.`;
      }
    }

    const selectedLeaveType = STATIC_LEAVE_TYPES.find(type => type.Id === formState.leaveTypeId);
    if (selectedLeaveType?.Title === 'Sick Leave' && formState.totalDays > 1 && formState.newAttachments.length === 0) {
      newErrors.attachment = 'Medical certificate required for sick leave exceeding 1 day';
    }

    for (const file of formState.newAttachments) {
      const isValidType = FILE_CONFIG.ACCEPTED_TYPES.includes(file.type as AcceptedFileType);
      if (!isValidType) newErrors.attachment = 'Only PDF, JPG, and PNG files are allowed';
      else if (file.size > FILE_CONFIG.MAX_SIZE) newErrors.attachment = 'File size must be less than 5MB';
    }

    return newErrors;
  };

  handleFieldChange = (field: keyof IFormState, value: any) => {
    this.setState(prevState => ({
      formState: { ...prevState.formState, [field]: value },
      errors: { ...prevState.errors, [field]: undefined }
    }), () => {
      if (field === 'leaveDurationType') {
        const isHalfDay = value === 'Half Day';
        this.setState({ showHalfDayType: isHalfDay });

        if (isHalfDay && this.state.formState.startDate) {
          this.setState(prevState => ({
            formState: { ...prevState.formState, endDate: prevState.formState.startDate }
          }), () => {
            this.updateTotalDays();
          });
        } else if (!isHalfDay) {
          const { startDate, endDate } = this.state.formState;
          if (startDate && endDate && new Date(startDate) > new Date(endDate)) {
            this.setState(prevState => ({
              formState: { ...prevState.formState, endDate: '' }
            }), () => {
              this.updateTotalDays();
            });
          } else {
            this.updateTotalDays();
          }
        }
      } else if (field === 'startDate') {
        const isHalfDay = this.state.formState.leaveDurationType === 'Half Day';

        if (isHalfDay) {
          this.setState(prevState => ({
            formState: { ...prevState.formState, endDate: value }
          }), () => {
            this.updateTotalDays();
          });
        } else {
          const { endDate } = this.state.formState;
          if (value && endDate && new Date(value) > new Date(endDate)) {
            this.setState(prevState => ({
              formState: { ...prevState.formState, endDate: '' }
            }), () => {
              this.updateTotalDays();
            });
          } else {
            this.updateTotalDays();
          }
        }
      } else if (field === 'endDate') {
        this.updateTotalDays();
      }
    });
  };

  // ✅ Handle multiple file selection
  handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const filesArray = Array.from(e.target.files);
      this.setState(prevState => ({
        formState: {
          ...prevState.formState,
          newAttachments: [...prevState.formState.newAttachments, ...filesArray]
        },
        errors: { ...prevState.errors, attachment: undefined }
      }));
    }
  };

  // ✅ Remove new attachment from list
  removeNewAttachment = (index: number) => {
    this.setState(prevState => ({
      formState: {
        ...prevState.formState,
        newAttachments: prevState.formState.newAttachments.filter((_, i) => i !== index)
      }
    }));
  };

  updateTotalDays = () => {
    const { startDate, endDate, leaveDurationType } = this.state.formState;

    if (leaveDurationType === 'Half Day' && startDate) {
      this.setState(prevState => ({
        formState: { ...prevState.formState, totalDays: 0.5 }
      }));
      return;
    }

    if (!startDate || !endDate) {
      this.setState(prevState => ({
        formState: { ...prevState.formState, totalDays: 0 }
      }));
      return;
    }

    const start = new Date(startDate);
    const end = new Date(endDate);
    if (isNaN(start.getTime()) || isNaN(end.getTime())) return;

    const weekendErrorMsg = validateWeekendDates(start, end);
    this.setState({ weekendError: weekendErrorMsg });

    if (!weekendErrorMsg && end >= start) {
      const days = calculateTotalDays(start, end, false);
      this.setState(prevState => ({
        formState: { ...prevState.formState, totalDays: days }
      }));
    } else {
      this.setState(prevState => ({
        formState: { ...prevState.formState, totalDays: 0 }
      }));
    }
  };

  handleSuccessClose = () => {
    this.setState({ showSuccessPopup: false });

    if (this.props.onLeaveSubmitted) {
      this.props.onLeaveSubmitted();
    }
  };

  redirectToDashboard = () => {
    this.setState({ showProbationPopup: false });
    if (this.props.onLeaveSubmitted) {
      this.props.onLeaveSubmitted();
    }
  };

  handleSubmit = () => {
    const { formState } = this.state;

    if (this.isProbation()) {
      this.setState({ showProbationPopup: true });
      return;
    }

    if (formState.startDate && formState.endDate) {
      const start = new Date(formState.startDate);
      const end = new Date(formState.endDate);
      const weekendErrorMsg = validateWeekendDates(start, end);
      if (weekendErrorMsg) {
        this.setState({ errors: { weekend: weekendErrorMsg } });
        return;
      }
    }

    const validationErrors = this.validateForm();
    if (Object.keys(validationErrors).length > 0) {
      this.setState({ errors: validationErrors });
      return;
    }

    if (!this.state.employee?.managerId) {
      this.setState({ errors: { manager: 'No manager assigned. Please contact HR.' } });
      return;
    }

    this.setState({ showConfirmationPopup: true });
  };

  // ✅ Upload multiple attachments
  uploadAttachments = async (requestId: number) => {
    const { formState } = this.state;
    
    for (const file of formState.newAttachments) {
      const buffer = await file.arrayBuffer();
      await sp.web.lists
        .getByTitle('Leave Requests')
        .items
        .getById(requestId)
        .attachmentFiles
        .add(file.name, buffer);
    }
  };

  confirmSubmit = async () => {
    const { formState, employee } = this.state;

    this.setState({ isSubmitting: true, showConfirmationPopup: false });

    try {
      const leaveRequestData: any = {
        Title: `Leave Request - ${employee?.name}`,
        EmployeeId: employee?.id,
        LineManagerId: employee?.managerId,
        DepartmentId: employee?.departmentId,
        CountryId: employee?.countryId,
        CountryCodeId: employee?.countryCodeId,
        LeaveTypeId: formState.leaveTypeId,
        OtherLeaveType: formState.otherLeaveType || null,
        LeaveDurationType: formState.leaveDurationType,
        HalfDayType: formState.halfDayType,
        StartDate: formState.startDate,
        EndDate: formState.endDate,
        TotalDays: formState.totalDays,
        Reason: formState.reason,
        Status: 'Pending on Manager',
        EmploymentType: employee?.employmentType,
        EmployeeYearCycle: this.state.currentCycle ? this.state.currentCycle.cycleNumber : 404
      };

      const result = await sp.web.lists.getByTitle('Leave Requests').items.add(leaveRequestData);

      // ✅ Upload multiple attachments
      if (formState.newAttachments.length > 0) {
        await this.uploadAttachments(result.data.Id);
      }

      if (employee && this.state.currentCycle) {
        const cacheKey = `leave_requests_${employee.id}_${this.state.currentCycle.cycleNumber}`;
        simpleCache.delete(cacheKey);
      }

      this.setState({
        formState: { ...FORM_INITIAL_STATE, leaveTypeId: 0, newAttachments: [] },
        errors: {},
        weekendError: null
      });

      if (employee && this.state.currentCycle) {
        const updatedRequests = await this.fetchLeaveRequestsFast(employee.id, this.state.currentCycle.cycleNumber);
        const updatedUsedLeaves = this.calculateUsedLeavesFast(updatedRequests);
        this.setState({ 
          existingRequests: updatedRequests,
          usedLeavesInCycle: updatedUsedLeaves
        });
      }

      this.setState({ showSuccessPopup: true });

    } catch (error: any) {
      if (this.isMounted) {
        this.setState({ errors: { submit: error.message || 'Error submitting leave request' } });
      }
    } finally {
      if (this.isMounted) {
        this.setState({ isSubmitting: false });
      }
    }
  };

  componentDidUpdate(prevProps: any, prevState: IApplyLeaveState) {
    if (prevState.formState.leaveDurationType !== this.state.formState.leaveDurationType &&
      this.state.formState.leaveDurationType === 'Half Day' &&
      this.state.formState.startDate) {
      this.handleFieldChange('endDate', this.state.formState.startDate);
    }
  }

  renderLoading = () => (
    <div className={styles.loadingContainer}>
      <div className={styles.spinner}></div>
      <p>Loading your leave information...</p>
    </div>
  );

  render() {
    const { showEditLeave, loading, employee, formState, errors, isSubmitting, showProbationPopup, showConfirmationPopup, showSuccessPopup, weekendError, showHalfDayType } = this.state;

    if (showEditLeave) {
      return <EditLeave />;
    }

    if (loading) {
      return this.renderLoading();
    }

    if (!employee) {
      return (
        <div className={styles.errorContainer}>
          <p>No employee data found. Please contact HR.</p>
          <button onClick={() => window.location.reload()}>Retry</button>
        </div>
      );
    }

    const showAttachmentSection = formState.leaveTypeId === 3 && formState.totalDays > 1;
    const hasAttachments = formState.newAttachments.length > 0;

    return (
      <>
        {/* Probation Popup */}
        {showProbationPopup && (
          <div className={styles.probationPopupOverlay}>
            <div className={styles.probationPopup}>
              <div className={styles.probationPopupHeader}>
                <h3>Probation Period Notice</h3>
              </div>
              <div className={styles.probationPopupBody}>
                <p>Dear <strong>{employee.name}</strong>,</p>
                <p>As per company policy, employees on <strong>Probation Period</strong> are <strong>not eligible</strong> to apply for any type of leave.</p>
                <p>Leave benefits will be available after successful completion of probation period.</p>
                <p style={{ marginTop: '16px', color: '#ef4444', fontWeight: '500' }}>
                  ⚠️ You cannot apply for leave during probation.
                </p>
              </div>
              <div className={styles.probationPopupFooter}>
                <button className={styles.probationPopupButton} onClick={this.redirectToDashboard}>
                  Got It
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Confirmation Popup */}
        {showConfirmationPopup && (
          <div className={styles.probationPopupOverlay} onClick={() => this.setState({ showConfirmationPopup: false })}>
            <div className={styles.probationPopup} onClick={e => e.stopPropagation()}>
              <div className={styles.probationPopupHeader}>
                <h3>Confirm Leave Request</h3>
              </div>
              <div className={styles.probationPopupBody}>
                <p>Are you sure you want to submit this leave request?</p>
                <p><strong>Leave Type:</strong> {STATIC_LEAVE_TYPES.find(t => t.Id === formState.leaveTypeId)?.Title}</p>
                <p><strong>Duration:</strong> {formState.totalDays} day(s)</p>
                <p><strong>Dates:</strong> {formState.startDate} to {formState.endDate}</p>
                <p><strong>Reason:</strong> {formState.reason}</p>
                {hasAttachments && (
                  <p><strong>Attachments:</strong> {formState.newAttachments.length} file(s)</p>
                )}
              </div>
              <div className={styles.probationPopupFooter} style={{ gap: '12px' }}>
                <button className={styles.probationPopupButton} style={{ background: '#94a3b8' }} onClick={() => this.setState({ showConfirmationPopup: false })}>Cancel</button>
                <button className={styles.probationPopupButton} onClick={this.confirmSubmit} disabled={isSubmitting}>
                  {isSubmitting ? 'Submitting...' : 'Confirm'}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Success Popup */}
        {showSuccessPopup && (
          <div className={styles.probationPopupOverlay} onClick={this.handleSuccessClose}>
            <div className={styles.probationPopup} onClick={e => e.stopPropagation()}>
              <div className={styles.probationPopupHeader} style={{ background: '#10b981' }}>
                <h3 style={{ color: 'white' }}>✓ Success!</h3>
              </div>
              <div className={styles.probationPopupBody}>
                <svg width="64" height="64" viewBox="0 0 24 24" fill="none" style={{ margin: '0 auto 16px', display: 'block' }}>
                  <path d="M20 6L9 17L4 12" stroke="#10b981" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
                <p style={{ fontSize: '16px', textAlign: 'center' }}>
                  <strong>Your leave request has been submitted successfully!</strong>
                </p>
                <p style={{ fontSize: '14px', color: '#64748b', textAlign: 'center', marginTop: '8px' }}>
                  Your request has been sent to the concern Department.
                </p>
              </div>
              <div className={styles.probationPopupFooter}>
                <button
                  className={styles.probationPopupButton}
                  style={{ background: '#10b981' }}
                  onClick={this.handleSuccessClose}
                >
                  Go to Dashboard
                </button>
              </div>
            </div>
          </div>
        )}

        <div className={styles.applyLeaveFormContent}>
          {(Object.values(errors).some(error => error) || weekendError) && (
            <div className={styles.errorSummary}>
              {weekendError && <div className={styles.errorMessage}>{weekendError}</div>}
              {Object.values(errors).map((error, idx) => error && <div key={idx} className={styles.errorMessage}>{error}</div>)}
            </div>
          )}

          <div className={styles.applyLeaveForm}>
            <div className={styles.applyLeaveRow}>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Employee Name</label>
                <input type="text" value={employee.name} disabled className={styles.applyLeaveDisabledField} />
              </div>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Department</label>
                <input type="text" value={employee.department} disabled className={styles.applyLeaveDisabledField} />
              </div>
            </div>

            <div className={styles.applyLeaveRow}>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Country</label>
                <input type="text" value={employee.country} disabled className={styles.applyLeaveDisabledField} />
              </div>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Manager</label>
                <input type="text" value={employee.managerName || 'Not Assigned'} disabled className={styles.applyLeaveDisabledField} />
              </div>
            </div>

            <div className={styles.applyLeaveRow}>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Leave Type <span className={styles.applyLeaveRequired}>*</span></label>
                <select 
                  value={formState.leaveTypeId} 
                  onChange={e => this.handleFieldChange('leaveTypeId', parseInt(e.target.value))} 
                  className={styles.applyLeaveSelect} 
                  disabled={this.isProbation()}
                >
                  <option value="0">Select Leave Type</option>
                  {STATIC_LEAVE_TYPES.map(type => (
                    <option key={type.Id} value={type.Id}>
                      {type.Title}
                    </option>
                  ))}
                </select>
                {this.isProbation() && (
                  <div className={styles.probationNote}>⚠️ You cannot apply for leave during probation period</div>
                )}
              </div>

              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Leave Duration <span className={styles.applyLeaveRequired}>*</span></label>
                <select 
                  value={formState.leaveDurationType} 
                  onChange={e => this.handleFieldChange('leaveDurationType', e.target.value as 'Full Day' | 'Half Day')} 
                  className={styles.applyLeaveSelect}
                  disabled={this.isProbation()}
                >
                  <option value="Full Day">Full Day</option>
                  <option value="Half Day">Half Day</option>
                </select>
              </div>
            </div>

            {formState.leaveTypeId === 4 && (
              <div className={styles.applyLeaveRow}>
                <div className={styles.applyLeaveFormGroup}>
                  <label className={styles.applyLeaveLabel}>Other Leave Type <span className={styles.applyLeaveRequired}>*</span></label>
                  <input 
                    type="text" 
                    value={formState.otherLeaveType} 
                    onChange={e => this.handleFieldChange('otherLeaveType', e.target.value)} 
                    placeholder="e.g., Emergency Leave, Marriage Leave, etc." 
                    className={styles.applyLeaveInput}
                    disabled={this.isProbation()}
                  />
                </div>
                {showHalfDayType && (
                  <div className={styles.applyLeaveFormGroup}>
                    <label className={styles.applyLeaveLabel}>Half Day Type <span className={styles.applyLeaveRequired}>*</span></label>
                    <select 
                      value={formState.halfDayType} 
                      onChange={e => this.handleFieldChange('halfDayType', e.target.value as 'First Half' | 'Second Half')} 
                      className={styles.applyLeaveSelect}
                      disabled={this.isProbation()}
                    >
                      <option value="First Half">First Half</option>
                      <option value="Second Half">Second Half</option>
                    </select>
                  </div>
                )}
              </div>
            )}

            {formState.leaveTypeId !== 4 && showHalfDayType && (
              <div className={styles.applyLeaveRow}>
                <div className={styles.applyLeaveFormGroup}>
                  <label className={styles.applyLeaveLabel}>Half Day Type <span className={styles.applyLeaveRequired}>*</span></label>
                  <select 
                    value={formState.halfDayType} 
                    onChange={e => this.handleFieldChange('halfDayType', e.target.value as 'First Half' | 'Second Half')} 
                    className={styles.applyLeaveSelect}
                    disabled={this.isProbation()}
                  >
                    <option value="First Half">First Half</option>
                    <option value="Second Half">Second Half</option>
                  </select>
                </div>
              </div>
            )}

            <div className={styles.applyLeaveRow}>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Start Date <span className={styles.applyLeaveRequired}>*</span></label>
                <input 
                  type="date" 
                  value={formState.startDate}
                  onChange={e => this.handleFieldChange('startDate', e.target.value)} 
                  className={styles.applyLeaveInput}
                  disabled={this.isProbation()}
                />
              </div>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>End Date {formState.leaveDurationType !== 'Half Day' && <span className={styles.applyLeaveRequired}>*</span>}</label>
                <input 
                  type="date" 
                  value={formState.endDate} 
                  disabled={formState.leaveDurationType === 'Half Day' || this.isProbation()} 
                  min={formState.startDate || formatDate(new Date())} 
                  onChange={e => this.handleFieldChange('endDate', e.target.value)} 
                  className={styles.applyLeaveInput} 
                />
              </div>
            </div>

            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>Total Days</label>
              <input 
                type="number" 
                value={formState.totalDays} 
                disabled 
                className={styles.applyLeaveTotalDaysField} 
                step="0.5" 
              />
            </div>

            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>Reason for Leave <span className={styles.applyLeaveRequired}>*</span></label>
              <textarea 
                value={formState.reason} 
                onChange={e => this.handleFieldChange('reason', e.target.value)} 
                placeholder="Please provide reason for your leave" 
                rows={4} 
                className={styles.applyLeaveTextarea}
                disabled={this.isProbation()}
              />
            </div>

            {/* ✅ Attachments Section - Only for Sick Leave */}
            {showAttachmentSection && (
              <>
                {/* Attachments Gallery */}
                {hasAttachments && (
                  <div className={styles.applyLeaveFormGroup}>
                    <label className={styles.applyLeaveLabel}>Medical Evidence</label>
                    <div className={styles.attachmentsGallery}>
                      {formState.newAttachments.map((file: File, index: number) => (
                        <div key={`new-${index}`} className={styles.applyLeaveFilePreview}>
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
                            <path d="M21 15V19C21 19.5304 20.7893 20.0391 20.4142 20.4142C20.0391 20.7893 19.5304 21 19 21H5C4.46957 21 3.96086 20.7893 3.58579 20.4142C3.21071 20.0391 3 19.5304 3 19V15M7 10L12 15M12 15L17 10M12 15V3" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                          </svg>
                          <span>{file.name}</span>
                          {!this.isProbation() && (
                            <span className={styles.applyLeaveRemoveFile} onClick={() => this.removeNewAttachment(index)}>
                              <svg width="14" height="14" viewBox="0 0 24 24" fill="none">
                                <path d="M18 6L6 18M6 6L18 18" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                              </svg>
                            </span>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* Upload New Files Button */}
                <div className={styles.applyLeaveFormGroup}>
                  <label className={styles.applyLeaveLabel}>
                    Upload Medical Evidence {!hasAttachments && <span className={styles.applyLeaveRequired}>*</span>}
                  </label>
                  <div
                    className={styles.applyLeaveFileDropZone}
                    onClick={() => !this.isProbation() && document.getElementById('applyLeaveFileInput')?.click()}
                    style={{ cursor: this.isProbation() ? 'not-allowed' : 'pointer', opacity: this.isProbation() ? 0.6 : 1 }}
                  >
                    <div className={styles.applyLeaveUploadPrompt}>
                      <svg width="32" height="32" viewBox="0 0 24 24" fill="none">
                        <path d="M12 16V4M12 4L8 8M12 4L16 8M20 16V19C20 19.5304 19.7893 20.0391 19.4142 20.4142C19.0391 20.7893 18.5304 21 18 21H6C4.46957 21 3.96086 20.7893 3.58579 20.4142C3.21071 20.0391 3 19.5304 3 19V16" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                      <span>Click to upload or drag & drop</span>
                      <span className={styles.applyLeaveFileHint}>Supported: PDF, JPG, PNG (Max 5MB each)</span>
                    </div>
                  </div>
                  <input 
                    id="applyLeaveFileInput" 
                    type="file" 
                    style={{ display: 'none' }} 
                    accept=".pdf,.jpg,.jpeg,.png" 
                    multiple 
                    onChange={this.handleFileSelect} 
                    disabled={this.isProbation()}
                  />
                </div>
              </>
            )}

            <div className={styles.applyLeaveButtonContainer}>
              <button 
                type="button" 
                className={styles.applyLeaveSubmitButton} 
                onClick={this.handleSubmit} 
                disabled={isSubmitting || this.isProbation()}
              >
                {isSubmitting ? 'Submitting...' : 'Submit Leave Request'}
              </button>
            </div>
          </div>
        </div>
      </>
    );
  }
}