
"use client";

import { useState, useEffect } from "react";
import { Navbar } from "@/components/navbar";
import { Footer } from "@/components/footer";
import { useCollection, useMemoFirebase, useFirestore } from "@/firebase";
import { collection, serverTimestamp } from "firebase/firestore";
import { addDocumentNonBlocking } from "@/firebase/non-blocking-updates";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { 
  Dialog, 
  DialogContent, 
  DialogHeader, 
  DialogTitle, 
  DialogTrigger,
  DialogDescription,
  DialogFooter
} from "@/components/ui/dialog";
import { Textarea } from "@/components/ui/textarea";
import { 
  Select, 
  SelectContent, 
  SelectItem, 
  SelectTrigger, 
  SelectValue 
} from "@/components/ui/select";
import { 
  UserPlus, 
  Database, 
  Search, 
  FileUp, 
  ClipboardList, 
  CircleCheck, 
  CircleAlert, 
  Download, 
  FileSpreadsheet,
  Upload,
  Milk
} from "lucide-react";
import { useToast } from "@/hooks/use-toast";
import { read, utils, writeFile } from 'xlsx';
import { Badge } from "@/components/ui/badge";

export default function FarmersPage() {
  const firestore = useFirestore();
  const { toast } = useToast();
  const [searchTerm, setSearchTerm] = useState("");
  const [isAdding, setIsAdding] = useState(false);
  const [newFarmer, setNewFarmer] = useState({ name: "", canNumber: "", bankAccountNumber: "", ifscCode: "", milkType: "COW" });
  const [importData, setImportData] = useState("");
  const [isImporting, setIsImporting] = useState(false);

  const farmersQuery = useMemoFirebase(() => {
    if (!firestore) return null;
    return collection(firestore, 'farmers');
  }, [firestore]);

  const { data: farmers, isLoading } = useCollection(farmersQuery);

  const filteredFarmers = farmers?.filter(f => 
    f.name.toLowerCase().includes(searchTerm.toLowerCase()) || 
    f.canNumber.includes(searchTerm)
  ).sort((a, b) => {
    const aNum = parseInt(a.canNumber);
    const bNum = parseInt(b.canNumber);
    if (isNaN(aNum) || isNaN(bNum)) return a.canNumber.localeCompare(b.canNumber);
    return aNum - bNum;
  });

  const handleAddFarmer = () => {
    if (!newFarmer.name || !newFarmer.canNumber) {
      toast({ title: "Error", description: "Name and Can Number are required.", variant: "destructive" });
      return;
    }

    if (!firestore) return;

    addDocumentNonBlocking(collection(firestore, 'farmers'), {
      ...newFarmer,
      canNumber: newFarmer.canNumber.toString().padStart(3, '0'),
      active: true,
      createdAt: serverTimestamp(),
    });

    setNewFarmer({ name: "", canNumber: "", bankAccountNumber: "", ifscCode: "", milkType: "COW" });
    setIsAdding(false);
    toast({ title: "Success", description: "Farmer added successfully." });
  };

  const processImportArray = (data: any[]) => {
    if (!firestore) return;
    let count = 0;
    data.forEach((item: any) => {
      const normalizedItem: any = {};
      Object.keys(item).forEach(key => {
        const normalizedKey = key.toLowerCase().replace(/\s+/g, '');
        normalizedItem[normalizedKey] = item[key];
      });

      const name = item.Name || item.name || normalizedItem.name || normalizedItem.farmername;
      const canNumber = item['Can Number'] || item.canNumber || normalizedItem.cannumber || normalizedItem.can;
      const bankAccountNumber = item['Bank Account Number'] || item['Account Number'] || item.bankAccountNumber || normalizedItem.bankaccountnumber || normalizedItem.accountnumber || normalizedItem.account;
      const ifscCode = item['IFSC Code'] || item.ifscCode || normalizedItem.ifsccode || normalizedItem.ifsc;
      const milkTypeRaw = item['Milk Type'] || item.milkType || normalizedItem.milktype || normalizedItem.type;
      
      let milkType = "COW";
      if (milkTypeRaw?.toString().toUpperCase().includes("BUFFALO")) milkType = "BUFFALO";

      if (name && canNumber) {
        addDocumentNonBlocking(collection(firestore, 'farmers'), {
          name: name.toString(),
          canNumber: canNumber.toString().padStart(3, '0'),
          bankAccountNumber: (bankAccountNumber || "").toString(),
          ifscCode: (ifscCode || "").toString(),
          milkType: milkType,
          active: true,
          createdAt: serverTimestamp(),
        });
        count++;
      }
    });

    toast({ title: "Import Successful", description: `Added ${count} farmers.` });
    setIsImporting(false);
    setImportData("");
  };

  const handleBulkImportText = () => {
    try {
      let data: any[] = [];
      const trimmedData = importData.trim();
      
      if (!trimmedData) throw new Error("Please paste your data first.");

      if (trimmedData.startsWith('[')) {
        data = JSON.parse(trimmedData);
      } else {
        const lines = trimmedData.split('\n');
        if (lines.length < 1) throw new Error("No data found.");

        const firstLine = lines[0];
        const delimiter = firstLine.includes('\t') ? '\t' : (firstLine.includes(',') ? ',' : null);
        
        if (!delimiter) {
          throw new Error("Could not detect delimiter. Please use Comma (CSV) or paste directly from Excel (Tabs).");
        }

        const headers = lines[0].split(delimiter).map(h => h.trim());
        
        data = lines.slice(1).filter(line => line.trim()).map(line => {
          const values = line.split(delimiter);
          const obj: any = {};
          headers.forEach((header, i) => {
            obj[header] = values[i]?.trim();
          });
          return obj;
        });
      }

      processImportArray(data);
    } catch (e: any) {
      toast({ title: "Import Failed", description: e.message || "Invalid format.", variant: "destructive" });
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target?.result;
        if (!data) return;
        const wb = read(data, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const jsonData = utils.sheet_to_json(ws);
        processImportArray(jsonData);
      } catch (error) {
        toast({ title: "File Error", description: "Could not read the Excel file.", variant: "destructive" });
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const downloadExcelTemplate = () => {
    const templateData = [
      { "Name": "Rajesh Kumar", "Can Number": "101", "Bank Account Number": "9876543210", "IFSC Code": "SBIN0001234", "Milk Type": "COW" },
      { "Name": "Suresh Singh", "Can Number": "102", "Bank Account Number": "1234567890", "IFSC Code": "HDFC0005678", "Milk Type": "BUFFALO" }
    ];
    const ws = utils.json_to_sheet(templateData);
    const wb = utils.book_new();
    utils.book_append_sheet(wb, ws, "Farmers Template");
    writeFile(wb, "Farmers_Import_Template.xlsx");
    
    toast({ title: "Template Downloaded", description: "Excel (.xlsx) file is ready." });
  };

  const seedData = () => {
    if (!firestore) return;
    toast({ title: "Seeding...", description: "Adding 50 sample farmers." });
    for (let i = 1; i <= 50; i++) {
      addDocumentNonBlocking(collection(firestore, 'farmers'), {
        name: `Farmer ${i}`,
        canNumber: i.toString().padStart(3, '0'),
        bankAccountNumber: `ACC-${i.toString().padStart(6, '0')}`,
        ifscCode: `IFSC${i.toString().padStart(4, '0')}`,
        milkType: i % 2 === 0 ? "BUFFALO" : "COW",
        active: true,
        createdAt: serverTimestamp(),
      });
    }
  };

  return (
    <div className="min-h-screen flex flex-col bg-background">
      <Navbar />
      <main className="flex-grow pt-24 pb-20">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-8">
            <div>
              <h1 className="text-3xl font-black text-primary tracking-tight">Farmer Management</h1>
              <p className="text-muted-foreground">Directory of suppliers and bank details.</p>
            </div>
            <div className="flex flex-wrap gap-2">
              <Button onClick={() => setIsAdding(!isAdding)} variant={isAdding ? "outline" : "default"} className="rounded-full">
                <UserPlus className="w-4 h-4 mr-2" />
                {isAdding ? "Cancel" : "Add Farmer"}
              </Button>
              
              <Dialog open={isImporting} onOpenChange={setIsImporting}>
                <DialogTrigger asChild>
                  <Button variant="secondary" className="rounded-full">
                    <FileUp className="w-4 h-4 mr-2" />
                    Bulk Import
                  </Button>
                </DialogTrigger>
                <DialogContent className="sm:max-w-[700px] rounded-3xl overflow-hidden p-0">
                  <div className="bg-primary p-6 text-primary-foreground">
                    <DialogHeader>
                      <DialogTitle className="text-2xl font-black flex items-center gap-2">
                        <FileSpreadsheet className="w-6 h-6" />
                        Excel Bulk Import
                      </DialogTitle>
                      <DialogDescription className="text-primary-foreground/80">
                        Upload your .xlsx file with Name, Can Number, Bank Account Number, IFSC Code, and Milk Type.
                      </DialogDescription>
                    </DialogHeader>
                  </div>
                  
                  <div className="p-6 space-y-6">
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div className="space-y-3">
                        <label className="text-xs font-black text-primary uppercase tracking-widest flex items-center gap-2">
                          <Upload className="w-3 h-3" /> Step 1: Upload File
                        </label>
                        <Input 
                          type="file" 
                          accept=".xlsx, .xls, .csv" 
                          onChange={handleFileUpload}
                          className="rounded-xl h-24 border-dashed border-2 cursor-pointer hover:bg-muted/50 transition-colors file:hidden text-center pt-8 text-muted-foreground font-medium"
                        />
                        <p className="text-[10px] text-center text-muted-foreground">Drop your Excel file here or click to browse</p>
                      </div>

                      <div className="space-y-3">
                        <div className="flex items-center justify-between">
                          <label className="text-xs font-black text-primary uppercase tracking-widest flex items-center gap-2">
                            <Download className="w-3 h-3" /> Step 2: Download Format
                          </label>
                        </div>
                        <div className="bg-muted/50 rounded-2xl p-4 border border-primary/10 flex flex-col items-center justify-center h-24 gap-2">
                          <Button variant="outline" size="sm" onClick={downloadExcelTemplate} className="rounded-full w-full bg-background shadow-sm hover:bg-primary hover:text-white transition-all">
                            <Download className="w-3 h-3 mr-2" /> Download Template (.xlsx)
                          </Button>
                          <p className="text-[10px] text-muted-foreground italic text-center">Includes Bank Account Number & IFSC</p>
                        </div>
                      </div>
                    </div>

                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <label className="text-sm font-bold text-primary flex items-center gap-2 uppercase tracking-wider">
                          <ClipboardList className="w-4 h-4" />
                          Or Paste Data
                        </label>
                      </div>
                      <Textarea 
                        placeholder="Name, Can Number, Bank Account Number, IFSC Code, Milk Type&#10;John Doe, 101, 9123456789, SBIN0001234, COW" 
                        value={importData}
                        onChange={(e) => setImportData(e.target.value)}
                        className="min-h-[120px] font-mono text-xs rounded-2xl border-primary/20 bg-background/50 focus:bg-background transition-colors"
                      />
                    </div>
                  </div>

                  <div className="bg-muted/50 p-6 flex flex-col sm:flex-row gap-4 justify-between items-center border-t">
                    <div className="flex items-center gap-2 text-xs text-muted-foreground font-medium italic">
                      <CircleAlert className="w-4 h-4 text-accent" />
                      Milk Type must be COW or BUFFALO
                    </div>
                    <div className="flex gap-2">
                      <Button variant="ghost" onClick={() => setIsImporting(false)} className="rounded-full">Cancel</Button>
                      <Button onClick={handleBulkImportText} className="rounded-full px-8 shadow-md" disabled={!importData}>
                        <CircleCheck className="w-4 h-4 mr-2" />
                        Process Text
                      </Button>
                    </div>
                  </div>
                </DialogContent>
              </Dialog>

              <Button variant="ghost" onClick={seedData} className="rounded-full text-muted-foreground border-dashed border-2 hover:bg-muted/50">
                <Database className="w-4 h-4 mr-2" />
                Seed Samples
              </Button>
            </div>
          </div>

          {isAdding && (
            <Card className="mb-8 border-primary/20 bg-primary/5 rounded-3xl shadow-sm animate-in fade-in slide-in-from-top-4 duration-300">
              <CardHeader>
                <CardTitle className="text-lg flex items-center gap-2">
                  <UserPlus className="w-5 h-5 text-primary" />
                  New Farmer Profile
                </CardTitle>
              </CardHeader>
              <CardContent>
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-6">
                  <div className="space-y-2">
                    <label className="text-xs font-black text-muted-foreground uppercase tracking-widest">Full Name</label>
                    <Input 
                      placeholder="e.g. John Doe" 
                      value={newFarmer.name}
                      onChange={(e) => setNewFarmer({...newFarmer, name: e.target.value})}
                      className="rounded-xl h-11 border-primary/10"
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-black text-muted-foreground uppercase tracking-widest">Can Number</label>
                    <Input 
                      placeholder="e.g. 101" 
                      value={newFarmer.canNumber}
                      onChange={(e) => setNewFarmer({...newFarmer, canNumber: e.target.value})}
                      className="rounded-xl h-11 border-primary/10"
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-black text-muted-foreground uppercase tracking-widest">Milk Type</label>
                    <Select value={newFarmer.milkType} onValueChange={(v) => setNewFarmer({...newFarmer, milkType: v})}>
                      <SelectTrigger className="rounded-xl h-11 border-primary/10">
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="COW">COW</SelectItem>
                        <SelectItem value="BUFFALO">BUFFALO</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-black text-muted-foreground uppercase tracking-widest">Bank Account Number</label>
                    <Input 
                      placeholder="e.g. 9123456789" 
                      value={newFarmer.bankAccountNumber}
                      onChange={(e) => setNewFarmer({...newFarmer, bankAccountNumber: e.target.value})}
                      className="rounded-xl h-11 border-primary/10"
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-black text-muted-foreground uppercase tracking-widest">IFSC Code</label>
                    <Input 
                      placeholder="e.g. SBIN0001234" 
                      value={newFarmer.ifscCode}
                      onChange={(e) => setNewFarmer({...newFarmer, ifscCode: e.target.value})}
                      className="rounded-xl h-11 border-primary/10"
                    />
                  </div>
                </div>
                <div className="mt-6 flex justify-end">
                  <Button onClick={handleAddFarmer} className="rounded-full px-8 shadow-md">Save Farmer</Button>
                </div>
              </CardContent>
            </Card>
          )}

          <div className="relative mb-6">
            <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-5 h-5 text-muted-foreground" />
            <Input 
              className="pl-12 h-14 bg-card rounded-2xl border-primary/10 shadow-sm focus:ring-primary/20 text-lg" 
              placeholder="Quick search by Name or Can Number..." 
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>

          <Card className="rounded-3xl overflow-hidden border-none shadow-xl bg-card/50 backdrop-blur-sm">
            <Table>
              <TableHeader className="bg-muted/50 border-b">
                <TableRow className="hover:bg-transparent">
                  <TableHead className="w-[100px] font-black text-primary pl-6 py-5">CAN</TableHead>
                  <TableHead className="font-black text-primary">Farmer Name</TableHead>
                  <TableHead className="w-[120px] font-black text-primary">Milk Type</TableHead>
                  <TableHead className="font-black text-primary">Bank Details</TableHead>
                  <TableHead className="text-right pr-6 font-black text-primary">Actions</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {isLoading ? (
                  <TableRow><TableCell colSpan={5} className="text-center py-20 text-muted-foreground animate-pulse font-medium">Loading farmer directory...</TableCell></TableRow>
                ) : filteredFarmers?.length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={5} className="text-center py-20">
                      <div className="flex flex-col items-center gap-3">
                        <div className="p-4 bg-muted rounded-full">
                          <ClipboardList className="w-10 h-10 text-muted-foreground/30" />
                        </div>
                        <p className="text-muted-foreground font-semibold">No farmers found matching your search.</p>
                      </div>
                    </TableCell>
                  </TableRow>
                ) : (
                  filteredFarmers?.map((farmer) => (
                    <TableRow key={farmer.id} className="group hover:bg-primary/5 transition-colors border-b last:border-0">
                      <TableCell className="font-black text-primary pl-6 text-lg">{farmer.canNumber}</TableCell>
                      <TableCell className="font-bold text-base text-foreground/80">{farmer.name}</TableCell>
                      <TableCell>
                        <Badge variant={farmer.milkType === 'BUFFALO' ? "secondary" : "outline"} className="rounded-full font-bold">
                          {farmer.milkType || 'COW'}
                        </Badge>
                      </TableCell>
                      <TableCell>
                        <div className="flex flex-col text-xs">
                          <span className="text-muted-foreground font-mono font-medium">{farmer.bankAccountNumber || "—"}</span>
                          {farmer.ifscCode && <span className="text-primary/60 font-black uppercase text-[9px] tracking-widest">{farmer.ifscCode}</span>}
                        </div>
                      </TableCell>
                      <TableCell className="text-right pr-6">
                        <Button variant="ghost" size="sm" className="rounded-full opacity-0 group-hover:opacity-100 transition-opacity font-bold text-primary">Edit Profile</Button>
                      </TableCell>
                    </TableRow>
                  ))
                )}
              </TableBody>
            </Table>
            <div className="p-5 bg-muted/20 text-center text-[10px] text-muted-foreground font-black uppercase tracking-[0.2em] border-t">
              Total Active Suppliers: {farmers?.length || 0}
            </div>
          </Card>
        </div>
      </main>
      <Footer />
    </div>
  );
}
import { db } from "@/lib/firebase"; // Adjust path to your firebase config
import { 
  collection, 
  getDocs, 
  doc, 
  deleteDoc, 
  writeBatch 
} from "firebase/firestore";

// --- Inside your Component ---

// 1. Delete a single farmer
const handleDeleteFarmer = async (farmerId: string) => {
  if (confirm("Are you sure you want to delete this farmer?")) {
    try {
      await deleteDoc(doc(db, "farmers", farmerId));
      alert("Farmer deleted successfully");
      // Optional: window.location.reload(); or update state to refresh list
    } catch (error) {
      console.error("Error deleting farmer:", error);
    }
  }
};

// 2. Delete ALL farmers
const handleDeleteAllFarmers = async () => {
  const password = prompt("Type 'DELETE ALL' to confirm deleting every farmer record:");
  
  if (password === "DELETE ALL") {
    try {
      const querySnapshot = await getDocs(collection(db, "farmers"));
      const batch = writeBatch(db);

      querySnapshot.forEach((document) => {
        batch.delete(document.ref);
      });

      await batch.commit();
      alert("All farmer records have been cleared.");
      window.location.reload();
    } catch (error) {
      console.error("Error deleting all farmers:", error);
    }
  }
};