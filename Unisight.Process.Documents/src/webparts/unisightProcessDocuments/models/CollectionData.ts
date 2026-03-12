export let _collectionData = [
    {
      uniqueId: "6e8f7d1d-fc50-4e2e-97d5-f31b21ee9977",
      DisplayName: "name",
      ColumnName: "name",
      minWidth: 50,
      maxWidth: 100,
      isLink: true
    },
    {
      uniqueId: "ceceb4fb-0e93-4c95-92d6-77d6ae13358a",
      DisplayName: "CreatedBy",
      ColumnName: "CreatedBy",
      minWidth: 50,
      maxWidth: 100,
      isLink: false
    },
    {
      uniqueId: "d2ddb5e8-9f84-4fdf-9de9-ed63d80fa3b4",
      DisplayName: "ModifiedBy",
      ColumnName: "ModifiedBy",
      minWidth: 50,
      maxWidth: 100,
      isLink: false
    },
    {
      uniqueId: "3a1b888d-18c5-4620-a8a9-0319bb4ea135",
      DisplayName: "Process",
      ColumnName: "owstaxIdMapp",
      minWidth: 50,
      maxWidth: 100,
      isLink: false
    }
  ];

  export let _collectionDataForTabs = [
    {
      uniqueId: "4ce20026-de37-4cf1-9bbc-a97781b2ed47",
      OrderOfTab: 1,
      TabName: "Beskriving",
      Source: "ProcessList",
      ShowItemCount: false,
      ShowIcon: false,
      Icon : "Info"
    },
    {
      uniqueId: "bb6e729e-f39c-4116-b613-f33314540c86",
      OrderOfTab: 2,
      TabName: "Documents",
      Source: "DocumentsSearch",
      ShowItemCount: true,
      ShowIcon: true,
      Icon: "DocumentSet"
    },
    {
      uniqueId: "7516f7da-fa08-4589-a6f9-2cd8d0dba62c",
      OrderOfTab: 3,
      TabName: "Pages",
      Source: "PagesSearch",
      ShowItemCount: true,
      ShowIcon: true,
      Icon: "FileASPX"
    },

  ];