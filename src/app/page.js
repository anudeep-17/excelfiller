import Excelreader from "@/Component/Excelwork/Excelreader";

export default function Home() {
  return (
    <div>
      <Excelreader excelsheet_name = "../Component/210HWFeedbackTemplate.xlsx" />
    </div>
  );
}
