import MsTeamsClient, { UPLOAD_CONCURRENCY } from "../msTeams";
import { RateLimitError } from "../rateLimit";
import { FileDetails } from "../types";

// Mock fetch
window.fetch = jest.fn();

describe("MsTeamsClient", () => {
  let client: MsTeamsClient;
  const mockIc3Token = "mock-ic3-token";
  const mockChatToken = "mock-chat-token";
  const mockPermissionsId = "mock-permissions-id";

  beforeEach(() => {
    client = new MsTeamsClient(mockIc3Token, mockChatToken, mockPermissionsId);
    (fetch as jest.Mock).mockClear();
  });

  describe("createObject", () => {
    it("should create an object and return its ID", async () => {
      const mockResponse = {
        id: "mock-document-id",
        headers: {
          get: jest.fn().mockReturnValue("mock-cv"),
        },
        ok: true,
        json: async () => ({ id: "mock-document-id" }),
      };
      (fetch as jest.Mock).mockResolvedValueOnce(mockResponse);

      // @ts-ignore - accessing private method for testing
      const result = await client.createObject();

      expect(result).toEqual({
        id: "mock-document-id",
        msCv: "mock-cv",
      });
      expect(fetch).toHaveBeenCalledWith(
        "https://as-prod.asyncgw.teams.microsoft.com/v1/objects/",
        expect.objectContaining({
          method: "POST",
          headers: expect.objectContaining({
            authorization: `Bearer ${mockIc3Token}`,
          }),
          body: expect.any(String),
        }),
      );
    });

    it("should throw an error if the request fails", async () => {
      (fetch as jest.Mock).mockResolvedValueOnce({
        ok: false,
        statusText: "Unauthorized",
      });

      // @ts-ignore - accessing private method for testing
      await expect(client.createObject()).rejects.toThrow(
        "Failed to create object: Unauthorized",
      );
    });
  });

  describe("uploadFileContent", () => {
    it("should upload file content successfully", async () => {
      (fetch as jest.Mock).mockResolvedValueOnce({
        ok: true,
      });

      const mockFile: FileDetails = {
        name: "test.png",
        size: 1024,
        type: "image/png",
        base64: "dGVzdA==", // test in base64
      };

      // @ts-ignore - accessing private method for testing
      await expect(
        client.uploadFileContent("mock-document-id", "mock-cv", mockFile),
      ).resolves.not.toThrow();

      expect(fetch).toHaveBeenCalledWith(
        "https://as-prod.asyncgw.teams.microsoft.com/v1/objects/mock-document-id/content/imgpsh",
        expect.objectContaining({
          method: "PUT",
          headers: expect.objectContaining({
            authorization: `Bearer ${mockIc3Token}`,
          }),
        }),
      );
    });

    it("should throw an error if the upload fails", async () => {
      const mockResponse = {
        ok: false,
        statusText: "Bad Request",
        text: async () => "Invalid file format",
      };
      (fetch as jest.Mock).mockResolvedValueOnce(mockResponse);

      const mockFile: FileDetails = {
        name: "test.png",
        size: 1024,
        type: "image/png",
        base64: "dGVzdA==", // test in base64
      };

      // @ts-ignore - accessing private method for testing
      await expect(
        client.uploadFileContent("mock-document-id", "mock-cv", mockFile),
      ).rejects.toThrow(
        "Failed to upload image: Bad Request - Invalid file format",
      );
    });
  });

  describe("sendMetadata", () => {
    it("should send metadata successfully", async () => {
      const mockResponse = {
        ok: true,
        json: async () => ({ success: true }),
      };
      (fetch as jest.Mock).mockResolvedValueOnce(mockResponse);

      // @ts-ignore - accessing private method for testing
      await expect(
        client.sendMetadata("mock-document-id", "mock-cv", ["test"]),
      ).resolves.not.toThrow();

      expect(fetch).toHaveBeenCalledWith(
        "https://teams.microsoft.com/api/csa/apac/api/v1/customemoji/metadata",
        expect.objectContaining({
          method: "POST",
          headers: expect.objectContaining({
            authorization: `Bearer ${mockChatToken}`,
          }),
          body: JSON.stringify({
            shortcuts: ["test"],
            documentId: "mock-document-id",
          }),
        }),
      );
    });

    it("should throw an error if sending metadata fails", async () => {
      (fetch as jest.Mock).mockResolvedValueOnce({
        ok: false,
        statusText: "Bad Request",
      });

      // @ts-ignore - accessing private method for testing
      await expect(
        client.sendMetadata("mock-document-id", "mock-cv", ["test"]),
      ).rejects.toThrow("Failed to create metadata: Bad Request");
    });

    it("should explain a 409 as a name that already exists", async () => {
      // HTTP/2 responses carry no status text
      (fetch as jest.Mock).mockResolvedValueOnce({
        ok: false,
        status: 409,
        statusText: "",
        text: async () => "",
      });

      // @ts-ignore - accessing private method for testing
      await expect(
        client.sendMetadata("mock-document-id", "mock-cv", ["test"]),
      ).rejects.toThrow(
        "An emoji named :test: probably already exists, so try another name (409 Conflict)",
      );
    });

    it("should include the status code and body when status text is empty", async () => {
      (fetch as jest.Mock).mockResolvedValueOnce({
        ok: false,
        status: 403,
        statusText: "",
        text: async () => '{"message":"Not allowed"}',
      });

      // @ts-ignore - accessing private method for testing
      await expect(
        client.sendMetadata("mock-document-id", "mock-cv", ["test"]),
      ).rejects.toThrow(
        'Failed to create metadata: 403 Forbidden - {"message":"Not allowed"}',
      );
    });
  });

  describe("uploadFile", () => {
    it("should upload a file successfully", async () => {
      // Mock the private methods
      const createObjectMock = jest
        .fn()
        .mockResolvedValue({ id: "mock-document-id", msCv: "mock-cv" });
      const uploadFileContentMock = jest.fn().mockResolvedValue(undefined);
      const sendMetadataMock = jest.fn().mockResolvedValue(undefined);

      // @ts-ignore - replacing private methods for testing
      client.createObject = createObjectMock;
      // @ts-ignore
      client.uploadFileContent = uploadFileContentMock;
      // @ts-ignore
      client.sendMetadata = sendMetadataMock;

      const mockFile: FileDetails = {
        name: "test.png",
        size: 1024,
        type: "image/png",
        base64: "dGVzdA==", // test in base64
      };

      await client.uploadFile(mockFile);

      expect(createObjectMock).toHaveBeenCalled();
      expect(uploadFileContentMock).toHaveBeenCalledWith(
        "mock-document-id",
        "mock-cv",
        mockFile,
      );
      expect(sendMetadataMock).toHaveBeenCalledWith(
        "mock-document-id",
        "mock-cv",
        ["test"],
      );
    });

    it("should use the custom shortcut when provided", async () => {
      // @ts-ignore - replacing private methods for testing
      client.createObject = jest
        .fn()
        .mockResolvedValue({ id: "mock-document-id", msCv: "mock-cv" });
      // @ts-ignore
      client.uploadFileContent = jest.fn().mockResolvedValue(undefined);
      const sendMetadataMock = jest.fn().mockResolvedValue(undefined);
      // @ts-ignore
      client.sendMetadata = sendMetadataMock;

      await client.uploadFile({
        name: "test.png",
        size: 1024,
        type: "image/png",
        base64: "dGVzdA==",
        shortcut: "party-parrot",
      });

      expect(sendMetadataMock).toHaveBeenCalledWith(
        "mock-document-id",
        "mock-cv",
        ["party-parrot"],
      );
    });

    it("should propagate errors during upload", async () => {
      const error = new Error("Upload failed");
      // Mock createObject to throw an error
      // @ts-ignore - replacing private method for testing
      client.createObject = jest.fn().mockRejectedValue(error);

      const mockFile: FileDetails = {
        name: "test.png",
        size: 1024,
        type: "image/png",
        base64: "dGVzdA==",
      };

      await expect(client.uploadFile(mockFile)).rejects.toThrow(
        "Upload failed",
      );
    });
  });

  describe("uploadFiles", () => {
    it("should return an error if no files are provided", async () => {
      const result = await client.uploadFiles([]);
      expect(result).toEqual({
        success: false,
        error: "Please select files first",
      });
    });

    it("should upload multiple files successfully", async () => {
      const uploadFileSpy = jest
        .spyOn(client, "uploadFile")
        .mockResolvedValue();

      const mockFiles: FileDetails[] = [
        {
          name: "test1.png",
          size: 1024,
          type: "image/png",
          base64: "dGVzdA==",
        },
        {
          name: "test2.png",
          size: 2048,
          type: "image/png",
          base64: "dGVzdA==",
        },
      ];

      const result = await client.uploadFiles(mockFiles);

      expect(uploadFileSpy).toHaveBeenCalledTimes(2);
      expect(uploadFileSpy).toHaveBeenCalledWith(
        mockFiles[0],
        expect.any(Function),
      );
      expect(uploadFileSpy).toHaveBeenCalledWith(
        mockFiles[1],
        expect.any(Function),
      );
      expect(result).toEqual({
        success: true,
        status: "Your emojis have been uploaded",
      });

      uploadFileSpy.mockRestore();
    });

    it("should report failed files and continue with the rest", async () => {
      const consoleSpy = jest
        .spyOn(console, "error")
        .mockImplementation(() => {});
      const uploadFileSpy = jest
        .spyOn(client, "uploadFile")
        .mockRejectedValueOnce(new Error("Upload failed"))
        .mockResolvedValueOnce();

      const mockFiles: FileDetails[] = [
        {
          name: "test1.png",
          size: 1024,
          type: "image/png",
          base64: "dGVzdA==",
        },
        {
          name: "test2.png",
          size: 2048,
          type: "image/png",
          base64: "dGVzdA==",
        },
      ];

      const result = await client.uploadFiles(mockFiles);

      expect(uploadFileSpy).toHaveBeenCalledTimes(2);
      expect(result).toEqual({
        success: false,
        error: "Failed to upload 1 of 2 emojis:\ntest1.png: Upload failed",
      });

      uploadFileSpy.mockRestore();
      consoleSpy.mockRestore();
    });

    it("should report progress for each file", async () => {
      jest.spyOn(console, "error").mockImplementation(() => {});
      // @ts-ignore - replacing private methods for testing
      client.createObject = jest
        .fn()
        .mockResolvedValue({ id: "mock-document-id", msCv: "mock-cv" });
      // @ts-ignore
      client.uploadFileContent = jest.fn(
        async (_id: string, _cv: string, file: FileDetails) => {
          if (file.name === "b.png") throw new Error("Bad image");
        },
      );
      // @ts-ignore
      client.sendMetadata = jest.fn().mockResolvedValue(undefined);

      const onProgress = jest.fn();
      await client.uploadFiles(
        [
          { name: "a.png", size: 1, type: "image/png", base64: "" },
          { name: "b.png", size: 1, type: "image/png", base64: "" },
        ],
        onProgress,
      );

      // Files upload concurrently, so check each file's own sequence
      const events = onProgress.mock.calls.map(([p]) => p);
      expect(events.filter((p) => p.index === 0)).toEqual([
        { index: 0, state: "uploading", step: 1 },
        { index: 0, state: "uploading", step: 2 },
        { index: 0, state: "uploading", step: 3 },
        { index: 0, state: "done", step: 3 },
      ]);
      expect(events.filter((p) => p.index === 1)).toEqual([
        { index: 1, state: "uploading", step: 1 },
        { index: 1, state: "uploading", step: 2 },
        { index: 1, state: "error", step: 0, error: "Bad image" },
      ]);
      jest.restoreAllMocks();
    });

    it("should upload at most UPLOAD_CONCURRENCY files at once", async () => {
      let active = 0;
      let maxActive = 0;
      const uploadFileSpy = jest
        .spyOn(client, "uploadFile")
        .mockImplementation(async () => {
          active++;
          maxActive = Math.max(maxActive, active);
          await new Promise((resolve) => setTimeout(resolve, 5));
          active--;
        });

      const files: FileDetails[] = Array.from(
        { length: UPLOAD_CONCURRENCY + 4 },
        (_, i) => ({
          name: `emoji${i}.png`,
          size: 1,
          type: "image/png",
          base64: "",
        }),
      );
      const result = await client.uploadFiles(files);

      expect(uploadFileSpy).toHaveBeenCalledTimes(UPLOAD_CONCURRENCY + 4);
      expect(maxActive).toBe(UPLOAD_CONCURRENCY);
      expect(result.success).toBe(true);
      uploadFileSpy.mockRestore();
    });

    it("should list failures in file order", async () => {
      jest.spyOn(console, "error").mockImplementation(() => {});
      // The first file fails last, after the others have finished
      jest
        .spyOn(client, "uploadFile")
        .mockImplementation(async (file: FileDetails) => {
          if (file.name === "a.png") {
            await new Promise((resolve) => setTimeout(resolve, 10));
            throw new Error("slow failure");
          }
          if (file.name === "c.png") throw new Error("fast failure");
        });

      const result = await client.uploadFiles(
        ["a.png", "b.png", "c.png"].map((name) => ({
          name,
          size: 1,
          type: "image/png",
          base64: "",
        })),
      );

      expect(result.error).toBe(
        "Failed to upload 2 of 3 emojis:\na.png: slow failure\nc.png: fast failure",
      );
      jest.restoreAllMocks();
    });
  });

  describe("custom emoji management", () => {
    const listUrl =
      "https://teams.microsoft.com/api/csa/apac/api/v1/customemoji/metadata";

    it("should list emoticons across categories", async () => {
      (fetch as jest.Mock).mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          continuationToken: "x",
          categories: [
            { id: "c1", title: "Custom Emoji", emoticons: [{ id: "a;1" }] },
            { id: "c2", title: "More", emoticons: [{ id: "b;2" }] },
          ],
        }),
      });

      const emojis = await client.listCustomEmojis();

      expect(emojis).toEqual([{ id: "a;1" }, { id: "b;2" }]);
      expect(fetch).toHaveBeenCalledWith(
        listUrl,
        expect.objectContaining({
          headers: expect.objectContaining({
            authorization: `Bearer ${mockChatToken}`,
          }),
        }),
      );
    });

    it("should throw a descriptive error when listing fails", async () => {
      (fetch as jest.Mock).mockResolvedValueOnce({
        ok: false,
        status: 401,
        statusText: "",
        text: async () => "",
      });

      await expect(client.listCustomEmojis()).rejects.toThrow(
        "Failed to load custom emojis: 401 Unauthorized",
      );
    });

    it("should delete by id, keeping the ; separator", async () => {
      (fetch as jest.Mock).mockResolvedValueOnce({ ok: true, status: 204 });

      await client.deleteCustomEmoji("deal-with-it;0-suk-d2-abc");

      expect(fetch).toHaveBeenCalledWith(
        `${listUrl}/deal-with-it;0-suk-d2-abc`,
        expect.objectContaining({
          method: "DELETE",
          headers: expect.objectContaining({
            authorization: `Bearer ${mockChatToken}`,
          }),
        }),
      );
    });

    it("should turn a 429 into a RateLimitError", async () => {
      (fetch as jest.Mock).mockResolvedValueOnce({
        ok: false,
        status: 429,
        statusText: "",
        headers: { get: () => null },
        text: async () =>
          '{"Throttled":true,"Dimension":"User","Limit":10,"WindowMinutes":1,"RetryAfterMinutes":1}',
      });

      const error = await client.deleteCustomEmoji("a;1").catch((e) => e);
      expect(error).toBeInstanceOf(RateLimitError);
      expect(error.message).toBe(
        "Teams allows 10 deletions per minute. Try again in 60s.",
      );
    });

    it("should explain a 403 when deleting", async () => {
      (fetch as jest.Mock).mockResolvedValueOnce({
        ok: false,
        status: 403,
        statusText: "",
        text: async () => "",
      });

      await expect(client.deleteCustomEmoji("a;1")).rejects.toThrow(
        "Not allowed to delete this emoji, probably because someone else added it (403 Forbidden)",
      );
    });

    it("should fall back to the next image view", async () => {
      const blob = new Blob(["gif"]);
      (fetch as jest.Mock)
        .mockResolvedValueOnce({
          ok: false,
          status: 404,
          statusText: "",
          text: async () => "",
        })
        .mockResolvedValueOnce({ ok: true, blob: async () => blob });

      await expect(client.fetchEmojiImage("0-wuk-d1-abc")).resolves.toBe(blob);
      expect((fetch as jest.Mock).mock.calls.map(([url]) => url)).toEqual([
        "https://as-prod.asyncgw.teams.microsoft.com/v1/objects/0-wuk-d1-abc/views/imgt2_anim",
        "https://as-prod.asyncgw.teams.microsoft.com/v1/objects/0-wuk-d1-abc/views/imgpsh_fullsize_anim",
      ]);
    });

    it("should throw when no image view loads", async () => {
      (fetch as jest.Mock).mockResolvedValue({
        ok: false,
        status: 404,
        statusText: "",
        text: async () => "",
      });

      await expect(client.fetchEmojiImage("x")).rejects.toThrow(
        "Failed to load emoji image: 404 Not Found",
      );
      (fetch as jest.Mock).mockReset();
    });
  });
});
