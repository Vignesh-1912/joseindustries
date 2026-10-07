import { useState } from "react";
import { Modal, Form, Input, message } from "antd";
import { api } from "../api/client";
import { Vendor } from "../types";
import { VendorGstinLookup } from "./GstinLookup";

interface QuickAddVendorModalProps {
  open: boolean;
  onClose: () => void;
  /** Called with the newly-created vendor right after a successful save, so the caller can select it and refresh its own vendor list. */
  onCreated: (vendor: Vendor) => void;
}

/**
 * A lightweight "add a vendor without leaving the document" modal, opened from
 * the small + button next to the Vendor picker on purchase bills. Mirrors
 * QuickAddCustomerModal; full details (address, etc.) can still be filled in
 * later from the Vendors page.
 */
export function QuickAddVendorModal({ open, onClose, onCreated }: QuickAddVendorModalProps) {
  const [form] = Form.useForm();
  const [saving, setSaving] = useState(false);

  async function handleSubmit() {
    const values = await form.validateFields();
    setSaving(true);
    try {
      const res = await api.post<{ vendor: Vendor }>("/vendors", values);
      message.success(`Vendor "${res.vendor.name}" added`);
      form.resetFields();
      onCreated(res.vendor);
    } catch (err) {
      message.error(err instanceof Error ? err.message : "Failed to add vendor");
    } finally {
      setSaving(false);
    }
  }

  function handleCancel() {
    form.resetFields();
    onClose();
  }

  return (
    <Modal
      title="Add New Vendor"
      open={open}
      onCancel={handleCancel}
      onOk={handleSubmit}
      confirmLoading={saving}
      destroyOnClose
      width={480}
    >
      <div style={{ marginBottom: 16 }}>
        <div style={{ marginBottom: 4 }}>Find by GSTIN</div>
        <VendorGstinLookup autoFocus onResult={(r) => { form.resetFields(); onCreated(r.vendor); }} />
        <div style={{ fontSize: 12, color: "#888", marginTop: 4 }}>
          An existing vendor is selected; a new GSTIN is fetched from the GST portal and added for you. Or fill in the details below.
        </div>
      </div>
      <Form form={form} layout="vertical" size="middle">
        <Form.Item name="name" label="Name" rules={[{ required: true, message: "Name is required" }]}>
          <Input placeholder="Vendor / company name" />
        </Form.Item>
        <Form.Item name="phone" label="Phone">
          <Input />
        </Form.Item>
        <Form.Item name="email" label="Email">
          <Input />
        </Form.Item>
        <Form.Item name="gstin" label="GSTIN">
          <Input />
        </Form.Item>
        <Form.Item name="state" label="State" extra="Used to work out CGST+SGST vs IGST on bills from this vendor.">
          <Input placeholder="e.g. Tamil Nadu" />
        </Form.Item>
      </Form>
    </Modal>
  );
}
